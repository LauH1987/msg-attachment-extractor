use byteorder::{ByteOrder, LittleEndian};
use ole;
use regex::Regex;
use std::collections::HashMap;
use std::io::{Read, Write};

fn main() {

    // TODO: Add command line arguments to set msg path
    let file =
        std::fs::File::open("/home/lau/IdeaProjects/Rust/msg-attachment-extractor/resources/unicode.msg")
            .unwrap();
    let parser = ole::Reader::new(file).unwrap();

    let attachment_entries = parser
        .iterate()
        .filter(|entry| entry.name().starts_with("__attach"));

    let attachment_children = attachment_entries.map(|e| e.children_nodes());

    let attachments = attachment_children
        .map(|v| {
            parser
                .iterate()
                .filter(|e| v.contains(&e.id()) && e.name().contains("_37"))
                .filter_map(|e| {
                    let re = Regex::new(r"^__.*\.0_(37..).*").unwrap();
                    let name = re.captures_iter(e.name()).next();
                    if let Some(capture) = name {
                        Some((capture[1].to_string(), e))
                    } else {
                        None
                    }
                })
                .collect::<HashMap<_, _>>()
        })
        .map(|e| {
            let short_filename = e.get("3704").map(|e| {
                let vec_u8 = read_entry_to_vec(&parser, *e);
                let vec_u16 = u8_to_16_vec(&vec_u8);
                String::from_utf16(&vec_u16).unwrap()
            });

            let long_filename = e.get("3707").map(|e| {
                let vec_u8 = read_entry_to_vec(&parser, *e);
                let vec_u16 = u8_to_16_vec(&vec_u8);
                String::from_utf16(&vec_u16).unwrap()
            });

            let data = e
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

        let mut extracted_file = std::fs::File::create(format!("./{}", filename)).unwrap();
        extracted_file.write_all(&a.data);
    }
}

struct Attachment {
    short_filename: Option<String>,
    long_filename: Option<String>,
    data: Vec<u8>,
}

fn read_entry_to_vec(parser: &ole::Reader, e: &ole::Entry) -> Vec<u8> {
    let slice = parser.get_entry_slice(e).unwrap();
    slice.bytes().collect::<Result<Vec<u8>, _>>().unwrap()
}

fn u8_to_16_vec(slice: &[u8]) -> Vec<u16>{
    slice
        .chunks_exact(2)
        .map(|e| LittleEndian::read_u16(e))
        .collect::<Vec<_>>()
}