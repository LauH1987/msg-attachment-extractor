use ole;
use std::io::{Read, Write};

fn main() {
    let file =
        std::fs::File::open("/home/lau/IdeaProjects/Rust/msg-extractor/resources/unicode.msg")
            .unwrap();
    let parser = ole::Reader::new(file).unwrap();
    // Iterate through the entries
    // for entry in parser.iterate().filter(|entry| entry.name().starts_with("__substg1.0_3701")) {
    //     println!("{}", entry);
    // }

    let mut fileEntries = parser
        .iterate()
        .filter(|entry| entry.name().starts_with("__substg1.0_3707"));
    // //We're going to extract a file from the OLE storage
    // let entry = fileEntries.skip(1).next().unwrap();
    // let mut slice = parser.get_entry_slice(entry).unwrap();
    // let mut buffer = std::vec::Vec::<u8>::with_capacity(slice.len());
    // slice.read_to_end(&mut buffer);
    // //Saves the extracted file
    // let mut extracted_file = std::fs::File::create("./fileTest.tif").unwrap();
    // extracted_file.write_all(&buffer[..]);

    let mut slice = parser
        .get_entry_slice(fileEntries.skip(1).next().unwrap())
        .unwrap();
    let mut buffer = std::vec::Vec::<u8>::with_capacity(slice.len());
    slice.read_to_end(&mut buffer);

    let s = String::from_utf16(&(buffer.iter().map(|&c| c as u16).collect::<Vec<_>>())).unwrap();
    let s = String::from_utf8(buffer).unwrap();
    println!("{}", s);
}

struct Attachment {
    short_filename: Option<String>,
    long_filename: Option<String>,
    data: Vec<u8>,
}
