#![allow(warnings)]
use anyhow::{Context, Result}; // Uniform error handling
use clap::{App, Arg};
use rusqlite::{params, Connection};
use std::path::Path;
use xlsxwriter::{Workbook, Worksheet};

fn write_table_to_worksheet(conn: &Connection, worksheet: &mut Worksheet, query: &str, headers: &[&str]) -> Result<()> {
    let mut stmt = conn.prepare(query)?;
    let rows = stmt.query_map(params![], |row| {
        (0..headers.len()).map(|i| {
            // Try to get as String, if that fails, try to get as i64 and then convert to String
            let result: rusqlite::Result<Option<String>> = row.get(i)
                .or_else(|_| row.get::<_, Option<i64>>(i).map(|opt| opt.map(|val| val.to_string())));
            result
        })
            .collect::<rusqlite::Result<Vec<Option<String>>>>()
    })?;

    // Write headers
    for (i, &header) in headers.iter().enumerate() {
        worksheet.write_string(0, i as u16, header, None)?;
    }

    // Write data rows, handling None as empty string
    for (row_idx, row_result) in rows.enumerate() {
        let row = row_result?;
        for (col_idx, cell) in row.iter().enumerate() {
            worksheet.write_string((row_idx + 1) as u32, col_idx as u16, &cell.clone().unwrap_or_default(), None)?;
        }
    }

    Ok(())
}

fn extract_data_to_excel(db_path: &Path, output_path: &Path) -> Result<()> {
    let conn = Connection::open(db_path).context("Failed to open database connection")?;

    // Correctly unwrap the Result from Workbook::new
    let workbook = Workbook::new(output_path.to_str().ok_or_else(|| anyhow::anyhow!("Failed to convert output path to string"))?).context("Failed to create workbook")?;

    let tables = vec![
        ("downloads", vec!["ID", "URL", "Target Path", "Start Time", "End Time", "Last Access Time", "Total Bytes", "Opened", "Referrer", "Tab URL", "Tab Referrer URL", "Mime Type", "Original Mime Type"],
         "SELECT downloads.id, url, target_path, \
                datetime(start_time/1000000-11644473600,'unixepoch') as start_time, \
                datetime(end_time/1000000-11644473600,'unixepoch') as end_time, \
                CASE WHEN last_access_time = 0 THEN 'N/A' ELSE datetime(last_access_time/1000000-11644473600,'unixepoch') END as last_access_time, \
                total_bytes, opened, referrer, tab_url, tab_referrer_url, mime_type, original_mime_type \
                FROM downloads \
                INNER JOIN downloads_url_chains ON downloads.id = downloads_url_chains.id \
                ORDER BY downloads.start_time DESC"),
        ("keyword_search_terms", vec!["Term", "URL", "Last Visit Time"],
         "SELECT keyword_search_terms.term, urls.url,
                CASE WHEN urls.last_visit_time = 0 THEN 'N/A' ELSE datetime(urls.last_visit_time/1000000-11644473600,'unixepoch') END as last_visit_time
                FROM keyword_search_terms
                INNER JOIN urls ON keyword_search_terms.url_id = urls.id
				ORDER BY last_visit_time DESC"),
        ("urls", vec!["URL", "Title", "Visit Count", "Last Visit Time", "Hidden"],
         "SELECT url, title, visit_count, \
                CASE WHEN last_visit_time = 0 THEN 'N/A' ELSE datetime(last_visit_time/1000000-11644473600,'unixepoch') END as last_visit_time, hidden \
                FROM urls"),
    ];

    for (table_name, headers, query) in tables {
        let mut worksheet = workbook.add_worksheet(Some(table_name)).context(format!("Failed to add {} worksheet", table_name))?;
        write_table_to_worksheet(&conn, &mut worksheet, query, &headers).context(format!("Failed to write data to {} worksheet", table_name))?;
    }

    // Close the workbook, properly handling the Result
    workbook.close().context("Failed to close the workbook")?;

    Ok(())
}

fn main() -> Result<()> {
    let matches = App::new("BrowserHistoryParser")
        .version("v1.0")
        .author("\nAuthor: Mohd Khairulazam")
        .about("Extracts data (table 'downloads', 'keyword_search_terms' & 'urls') from Chromium-based browsers' SQLite database into an Excel file.")
        .arg(Arg::new("filename")
            .short('f')
            .long("filename")
            .takes_value(true)
            .required(true)
            .help("SQLite Database file/path"))
        .arg(Arg::new("output")
            .short('o')
            .long("output")
            .takes_value(true)
            .required(true)
            .help("Output Excel file path"))
        .get_matches();

    let db_path = Path::new(matches.value_of("filename").unwrap());
    let output_path = Path::new(matches.value_of("output").unwrap());

    extract_data_to_excel(db_path, output_path).context("Failed to extract data to Excel")?;

    println!("\nDone!");

    Ok(())
}
