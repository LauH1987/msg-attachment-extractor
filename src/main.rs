use anyhow::{Context, Result};
use byteorder::{ByteOrder, LittleEndian};
use lazy_static::lazy_static;
use ole::{Entry, Reader};
use regex::Regex;
use std::collections::HashMap;
use std::fs::OpenOptions;
use std::io::{BufWriter, Read, Write};
use std::{
    borrow::Cow,
    env,
    fs::create_dir_all,
    fs::File,
    path::{Path, PathBuf},
};
use structopt::StructOpt;

fn main() -> Result<()> {
    let options = Options::from_args();
    let file = File::open(&options.msg_file)?;
    let parser = Reader::new(file)?;

    let attachments = get_attachments(&parser);

    let dir = get_or_create_dir(&options)?;

    for a in attachments {
        a?.write_to_file(&options, &dir)?;
    }

    Ok(())
}

fn get_attachments<'a>(parser: &'a Reader) -> impl Iterator<Item = Result<Attachment>> + 'a {
    let id_map: HashMap<u32, &Entry> = parser.iterate().map(|entry| (entry.id(), entry)).collect();

    let attachment_entries = parser
        .iterate()
        .filter(|entry| entry.name().starts_with("__attach"));

    let attachment_children = attachment_entries.map(Entry::children_nodes);

    attachment_children
        .map(move |att_children| children_to_att_code_map(&id_map, att_children))
        .map(move |map| {
            let short_filename: Option<String> = map
                .get("3704")
                .map::<Result<String>, _>(|e| {
                    let vec_u8 = read_entry_to_vec(&parser, *e)?;
                    let vec_u16 = u8_to_16_vec(&vec_u8);
                    String::from_utf16(&vec_u16).map_err(From::from)
                })
                .transpose()?;

            let long_filename: Option<String> = map
                .get("3707")
                .map::<Result<String>, _>(|e| {
                    let vec_u8 = read_entry_to_vec(&parser, *e)?;
                    let vec_u16 = u8_to_16_vec(&vec_u8);
                    String::from_utf16(&vec_u16).map_err(From::from)
                })
                .transpose()?;

            let data = map
                .get("3701")
                .map(|e| read_entry_to_vec(&parser, *e))
                .transpose()?
                .unwrap();

            Ok(Attachment {
                short_filename,
                long_filename,
                data,
            })
        })
}

/// Takes a list of children of an attachment Entry and returns a hashmap where each child
/// is mapped to a it's corresponding attachment property type code (see http://www.fileformat.info/format/outlookmsg/).
/// For attachments in an msg-file these all start with 0_37
fn children_to_att_code_map<'a>(
    id_map: &HashMap<u32, &'a Entry>,
    att_children: &[u32],
) -> HashMap<String, &'a Entry> {
    att_children
        .iter()
        .map(|id| id_map.get(id).unwrap())
        .filter_map(|&entry| extract_attachment_code(entry).map(|code| (code, entry)))
        .collect::<HashMap<_, _>>()
}

/// Extracts the 4 digit attachment property code starting with 37
/// if no 37-code is matched returns None
fn extract_attachment_code(e: &Entry) -> Option<String> {
    lazy_static! {
        static ref RE: Regex = Regex::new(r"^__.*\.0_(37..).*").unwrap();
    }
    let name = RE.captures_iter(e.name()).next();
    name.map(|capture| capture[1].to_string())
}

struct Attachment {
    short_filename: Option<String>,
    long_filename: Option<String>,
    data: Vec<u8>,
}

impl Attachment {
    fn write_to_file<P: AsRef<Path>>(&self, options: &Options, dir: P) -> Result<()> {
        let filename: &str = self.long_filename.as_ref().unwrap_or_else(|| {
            self.short_filename
                .as_ref()
                .expect("No long or short filename for attachment")
        });

        let prefix: Option<Cow<str>> = if options.prefix_filename {
            Some(options.msg_file.file_name().unwrap().to_string_lossy())
        } else {
            None
        };

        let filename: Cow<str> = match prefix {
            Some(prefix) => format!("{} {}", prefix, filename).into(),
            None => filename.into(),
        };

        let err_context = || format!("Failed to write file: {}", filename);

        let file = if options.overwrite {
            File::create(dir.as_ref().join(filename.as_ref())).with_context(err_context)?
        } else {
            create_unique_file(dir, &filename, None)
        };

        let mut extracted_file = BufWriter::new(file);

        extracted_file
            .write_all(&self.data)
            .with_context(err_context)?;

        extracted_file
            .flush()
            .with_context(err_context)
            .map_err(From::from)
    }
}

fn create_unique_file<P: AsRef<Path>>(dir: P, filename: &str, num: Option<u32>) -> File {
    let filename_new: Cow<str> = match num {
        None => filename.into(),
        Some(i) => {
            let filename: &Path = filename.as_ref();
            let stem = filename
                .file_stem()
                .map(|s| s.to_string_lossy())
                .unwrap_or_else(|| "".into());
            let ext = filename
                .extension()
                .map(|s| s.to_string_lossy())
                .unwrap_or_else(|| "".into());
            format!("{}_{}.{}", stem, i, ext).into()
        }
    };

    let file = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(dir.as_ref().join(filename_new.as_ref()));

    match file {
        Ok(file) => file,
        Err(_) => create_unique_file(dir, filename, Some(num.unwrap_or(0) + 1)),
    }
}

/// Create and return subdirectory if option is on else return current dir
fn get_or_create_dir(options: &Options) -> Result<PathBuf> {
    if options.subfolder {
        let mut dir = env::current_dir()?;
        let msg_name_stem = options
            .msg_file
            .file_stem()
            .unwrap()
            .to_string_lossy()
            .into_owned();
        dir.push(msg_name_stem);
        create_dir_all(&dir)?;
        Ok(dir)
    } else {
        env::current_dir().map_err(From::from)
    }
}

fn read_entry_to_vec(parser: &Reader, e: &Entry) -> Result<Vec<u8>> {
    let slice = parser.get_entry_slice(e)?;
    slice
        .bytes()
        .collect::<Result<Vec<u8>, _>>()
        .map_err(From::from)
}

fn u8_to_16_vec(slice: &[u8]) -> Vec<u16> {
    slice
        .chunks_exact(2)
        .map(|e| LittleEndian::read_u16(e))
        .collect::<Vec<_>>()
}

#[derive(StructOpt, Debug)]
#[structopt(name = "msg-attachment-extractor")]
struct Options {
    /// Prefix attachment filename with name of the msg-file
    #[structopt(long = "prefix")]
    prefix_filename: bool,

    /// Put extracted attachment in a subfolder with the name of the msg-file
    #[structopt(long)]
    subfolder: bool,

    /// Overwrite extracted files of same filename
    #[structopt(long)]
    overwrite: bool,

    /// File to process
    #[structopt(parse(from_os_str))]
    msg_file: PathBuf,
}
