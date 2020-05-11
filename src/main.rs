use byteorder::{ByteOrder, LittleEndian};
use lazy_static::lazy_static;
use ole;
use ole::{Entry, Reader};
use regex::Regex;
use std::collections::HashMap;
use std::io::{Read, Write};
use std::{fs::File, path::PathBuf};
use structopt::StructOpt;

fn main() {
    let options = Options::from_args();
    let file = File::open(options.msg_file).unwrap();
    let parser = Reader::new(file).unwrap();

    let attachment_entries = parser
        .iterate()
        .filter(|entry| entry.name().starts_with("__attach"));

    let attachment_children = attachment_entries.map(|e| e.children_nodes());

    let attachments = attachment_children
        .map(|att_children| children_to_att_code_map(&parser, att_children)) // TODO: This design is inefficient as is does multiple passes over the entire file.
                                                                             //       Maybe improve by making one hashmap with multiple keys using multi-map: One key for IDs and one key for property type code
        .map(|map| {
            let short_filename = map.get("3704").map(|e| {
                let vec_u8 = read_entry_to_vec(&parser, *e);
                let vec_u16 = u8_to_16_vec(&vec_u8);
                String::from_utf16(&vec_u16).unwrap()
            });

            let long_filename = map.get("3707").map(|e| {
                let vec_u8 = read_entry_to_vec(&parser, *e);
                let vec_u16 = u8_to_16_vec(&vec_u8);
                String::from_utf16(&vec_u16).unwrap()
            });

            let data = map
                .get("3701")
                .map(|e| read_entry_to_vec(&parser, *e))
                .unwrap();

            Attachment {
                short_filename,
                long_filename,
                data,
            }
        });

    for a in attachments {
        let filename: &str = a
            .long_filename
            .as_ref()
            .unwrap_or_else(|| a.short_filename.as_ref().unwrap());

        let mut extracted_file = File::create(format!("./{}", filename)).unwrap();
        extracted_file.write_all(&a.data);
    }
}

/// Takes a list of children of an attachment Entry and returns a hashmap where each child
/// is mapped to a it's corresponding attachment property type code (see http://www.fileformat.info/format/outlookmsg/).
/// For attachments in an msg-file these all start with 0_37
fn children_to_att_code_map<'a>(
    parser: &'a Reader,
    att_children: &[u32],
) -> HashMap<String, &'a Entry> {
    parser
        .iterate()
        .filter(|entry| att_children.contains(&entry.id()) && entry.name().contains("_37"))
        .filter_map(|entry| extract_attachment_code(entry).map(|code| (code, entry)))
        .collect::<HashMap<_, _>>()
}

/// Extracts the 4 digit attachment property code starting with 37
/// if no 37-code is matched returns None
fn extract_attachment_code(e: &Entry) -> Option<String> {
    lazy_static! {
        static ref RE: Regex = Regex::new(r"^__.*\.0_(37..).*").unwrap();
    }
    let name = RE.captures_iter(e.name()).next();
    if let Some(capture) = name {
        Some(capture[1].to_string())
    } else {
        None
    }
}

struct Attachment {
    short_filename: Option<String>,
    long_filename: Option<String>,
    data: Vec<u8>,
}

fn read_entry_to_vec(parser: &Reader, e: &Entry) -> Vec<u8> {
    let slice = parser.get_entry_slice(e).unwrap();
    slice.bytes().collect::<Result<Vec<u8>, _>>().unwrap()
}

fn u8_to_16_vec(slice: &[u8]) -> Vec<u16> {
    slice
        .chunks_exact(2)
        .map(|e| LittleEndian::read_u16(e))
        .collect::<Vec<_>>()
}

#[derive(StructOpt, Debug)]
#[structopt(name = "Options")]
struct Options {
    /// File to process
    #[structopt(parse(from_os_str))]
    msg_file: PathBuf,
}
