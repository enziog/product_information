extern crate calamine;
use std::io;
use std::fmt::Write;
use calamine::{Reader, Xlsx, open_workbook, DataType};

fn search(pat: &str) {
    let path = "S:\\20.Quality\\21 Project Information 项目信息平台\\竣工文件信息台账.xlsx";
    //let path = "C:\\Users\\jianhao.guo\\Desktop\\Equipment List.xlsx";
    let mut wb: Xlsx<_> = open_workbook(path).expect("Excel文件（xlsx格式）打开失败");
    let sheets = wb.sheet_names().to_owned();
    for sheet in sheets {
        if sheet.find("项目数据") != None { //查找sheet
            println!("iter_sheet: {}", sheet);
            if let Some(Ok(f)) = wb.worksheet_range(&sheet) {
                for row in f.rows() {
                    for (i, c) in row.iter().enumerate() {
                        if i == 6 {
                            let mut dest = String::new();
                                match *c {
                                    DataType::Empty => Ok(()),
                                    DataType::String(ref s) => write!(dest, "{}", s),
                                    DataType::Float(ref f) => write!(dest, "{}", f),
                                    DataType::Int(ref i) => write!(dest, "{}", i),
                                    DataType::Error(ref e) => write!(dest, "{:?}", e),
                                    DataType::Bool(ref b) => write!(dest, "{}", b),
                            };
                            println!("{}", dest);
                            //break;
                        }
                    }
                }
                //break;
            } else {
                continue;
            }
        }
    }
}

fn run() {
    let mut job_no= String::new();
    io::stdin().read_line(&mut job_no);
    search(&job_no);
    println!("按回车退出，输入工作令继续查找");
}
fn main() {
    run();
}
