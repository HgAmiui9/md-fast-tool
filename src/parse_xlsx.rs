use crate::consts::*;
use calamine::{open_workbook, Reader, Xlsx};

pub fn parse_xlsx() -> String {
    let mut excel: Xlsx<_> = open_workbook(FL_IOT).unwrap();
    let mut s = String::new();
    if let Some(Ok(r)) = excel.worksheet_range("工作表1") {
        let mut row_index = 1;
        for row in r.rows() {
            if row_index == 2 {
                s.push_str("|");
                for _ in 0..row.len() {
                    s.push_str(" --- ");
                    s.push_str("|");
                }
                s.push_str("\n");
            }

            s.push_str("|");
            for i in 0..row.len() {
                let row_str = &row[i].to_string().replace("\n", "");
                s.push_str(&format!(" {} ", row_str));
                s.push_str("|");
            }
            s.push_str("\n");
            row_index = row_index + 1;
        }
    }
    s
}
