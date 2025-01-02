use std::fs::File;
use std::io::{BufReader, BufRead};
use std::time::SystemTime;
use xlsxwriter::*;
use csv::ReaderBuilder;
use indicatif::ProgressBar;
use std::error::Error;
use chrono::Local;
use std::env;
use std::path::Path;

// 转换速度：
// python：547万，转换用时，680s（实际上会多等个几分钟），大约15分钟多
// rust：  547万，转换用时，87s，质的提升，2分钟以内
// rust: 4300万，开始：14:03:16，结束：预计20分钟以内，但是有可能因内存不够而导致电脑崩溃重启
// rust：20万 11s左右

// 程序功能：将 一个CSV 文件转换为 一个Excel 文件，并按指定行数分割为多个工作表，
// 当前配置为：每个工作表最多100万行，每次读取10万行

// 用法：
// 1. 将需要转换的CSV文件放在当前exe所在目录下，并且文件名必须为“a.csv”
// 2. 双击exe，运行程序，程序会自动读取CSV文件，并转换为Excel文件，并按指定行数分割为多个工作表
// 3. 转换后的Excel文件会保存在当前exe所在目录下，并且文件名为“a.xlsx”

// 配置文件：config.txt

// TODO: 现在是所有内容在最后一起写入，如果文件很大，有可能因内存不够而导致电脑崩溃重启

// 配置参数
// const ROWS_PER_SHEET: usize = 1_000_000;  // 每个工作表的最大行数
// const CHUNK_SIZE: usize = 100_000;  // 每次读取的行数，减小以降低内存使用

// 打包成release的exe：cargo b -r
// .\target\release\big-data-transfer.exe
fn main() -> Result<(), Box<dyn Error>> {
    // 获取当前 exe 文件所在目录
    let current_dir = env::current_exe()?.parent().unwrap().to_path_buf();

    // 定义配置文件路径
    let config_file_path = current_dir.join("config.txt");

    let config_file_path = config_file_path
        .to_str()
        .ok_or("目标文件路径包含非 UTF-8 字符")?;

    println!("当前配置文件路径: {}", config_file_path);

    // 检查配置文件是否存在
    if !Path::new(config_file_path).exists() {
        eprintln!("当前配置文件 {} 不存在！", config_file_path);
        return Err("配置文件不存在".into());
    }

    // 从配置文件中读取值
    let (rows_per_sheet, chunk_size) = read_config(config_file_path)?;
    println!("每个工作表的最大行数: {}", rows_per_sheet);
    println!("每次读取的行数: {}", chunk_size);

    // 定义输入和输出文件名
    let input_file = current_dir.join("a.csv");
    let output_file = current_dir.join("a.xlsx");

    // 检查输入文件是否存在
    if !input_file.exists() {
        eprintln!("源文件不存在：{}", input_file.display());
        return Err("源文件不存在".into());
    }

    let source_path = input_file
        .to_str()
        .ok_or("源文件路径包含非 UTF-8 字符")?;
    println!("源文件路径：{}", source_path);

    let output_prefix = output_file
        .to_str()
        .ok_or("目标文件路径包含非 UTF-8 字符")?;
    println!("目标文件路径：{}", output_prefix);

    let start_time = SystemTime::now();
    let formatted_date = Local::now().format("%Y-%m-%d %H:%M:%S").to_string();
    println!("开始时间：{}", formatted_date);

    // 计算CSV文件的总行数（不包括表头）
    let total_rows = count_csv_rows(source_path)?;
    println!("总行数（不含表头）：{}", total_rows);

    // 计算需要的工作表数量
    let num_sheets = (total_rows as f64 / rows_per_sheet as f64).ceil() as usize;
    println!("最终将生成的工作表数量：{}", num_sheets);

    
    let mut sheet_number = 1;
    let mut row_count = 0;

    // 创建进度条
    let pb = ProgressBar::new(total_rows as u64);

    // 逐块读取CSV文件
    let mut reader = ReaderBuilder::new()
        .has_headers(true)
        .from_reader(BufReader::new(File::open(source_path)?));

    // 提取表头
    let headers = reader.headers()?.clone(); // 获取表头并克隆

    // 初始化第一个工作表
    // let mut sheet = workbook.add_worksheet(Some(&format!("Sheet{}", sheet_number)))?;
    // let mut sheet: Option<Worksheet> = None;
    // sheet = Some(workbook.add_worksheet(Some(&format!("Sheet{}", sheet_number)))?);
    // let mut sheet = workbook.add_worksheet(Some(&format!("Sheet{}", sheet_number)))?;

    // // 写入表头到第一个工作表
    // for (j, value) in headers.iter().enumerate() {
    //     sheet.write_string(row_count as u32, j as u16, value, None)?;
    // }
    // row_count += 1;

    // 创建Excel文件，使用可变的 Workbook
    let workbook = Workbook::new(output_prefix)?;
    let mut sheet: Option<Worksheet> = None;

    // 初始化第一个工作表
    // sheet = Some(workbook.add_worksheet(Some(&format!("Sheet{}", sheet_number)))?);
    sheet = Some(workbook.add_worksheet(Some(&format!("Sheet{}", sheet_number)))?);

    // 写入表头到第一个工作表
    if let Some(ref mut s) = sheet {
        for (j, value) in headers.iter().enumerate() {
            s.write_string(row_count as u32, j as u16, value, None)?;
        }
    }
    row_count += 1;

    let mut chunk: Vec<csv::StringRecord> = Vec::with_capacity(chunk_size);

    for result in reader.records() {
        let record = result?;
        chunk.push(record);

        if chunk.len() >= chunk_size {
            
            write_chunk(&mut chunk, &mut sheet, &mut sheet_number, &headers, &mut row_count, rows_per_sheet, &workbook)?;

            pb.inc(chunk.len() as u64);
            chunk.clear();
        }
    }

    // 处理剩余的chunk
    if !chunk.is_empty() {

        // write_chunk(&mut chunk, &mut sheet, &mut sheet_number, &headers, &mut row_count, rows_per_sheet, &mut workbook)?;
        write_chunk(&mut chunk, &mut sheet, &mut sheet_number, &headers, &mut row_count, rows_per_sheet, &workbook)?;

        pb.inc(chunk.len() as u64);
    }

    pb.finish_with_message("CSV 处理完成");

    // 显式关闭 Workbook，确保所有数据写入文件
    workbook.close()?;  // 这里会触发最终的写入操作

    let elapsed_time = SystemTime::now().duration_since(start_time)?.as_secs();
    println!("成功将CSV转换为Excel文件：{}", output_prefix);
    println!("转换用时：{}s", elapsed_time);

    Ok(())
}

// write_chunk(&mut chunk, &mut sheet, &mut sheet_number, &headers, &mut row_count, rows_per_sheet, &mut workbook)?;

fn write_chunk<'a>(
    chunk: &mut Vec<csv::StringRecord>,
    sheet: &mut Option<Worksheet<'a>>, // 使用 Option 包装 Worksheet
    sheet_number: &mut usize,
    headers: &csv::StringRecord,
    row_count: &mut usize,
    rows_per_sheet: usize,
    workbook: &'a Workbook,
) -> Result<(), Box<dyn Error>> {
    for record in chunk.iter() {
        if *row_count >= rows_per_sheet + 1 { // +1 for header
            *sheet_number += 1;

            // 创建新工作表并重新赋值
            let new_sheet = workbook.add_worksheet(Some(&format!("Sheet{}", sheet_number)))?;
            *sheet = Some(new_sheet); // 更新 sheet 的值为新的工作表

            // 写入表头到新工作表
            if let Some(ref mut s) = sheet {
                for (j, value) in headers.iter().enumerate() {
                    s.write_string(0, j as u16, value, None)?;
                }
            }
            *row_count = 1; // 重置行数计数器
        }

        // 写入数据行
        if let Some(ref mut s) = sheet {
            for (j, value) in record.iter().enumerate() {
                s.write_string(*row_count as u32, j as u16, value, None)?;
            }
        }
        *row_count += 1;
    }

    chunk.clear();
    Ok(())
}

// fn write_chunk<'a>(
//     chunk: &mut Vec<csv::StringRecord>,
//     sheet: &mut Option<Worksheet<'a>>,  // 引用 Workbook 内的 Worksheet
//     sheet_number: &mut usize,
//     headers: &csv::StringRecord,
//     row_count: &mut usize,
//     rows_per_sheet: usize,
//     workbook: &'a mut Workbook,        // 显式声明 workbook 的生命周期
// ) -> Result<(), Box<dyn Error>> {
//     for record in chunk.iter() {
//         if *row_count >= rows_per_sheet + 1 { // +1 for header
//             *sheet_number += 1;

//             // 创建新工作表并重新赋值
//             let new_sheet = workbook.add_worksheet(Some(&format!("Sheet{}", sheet_number)))?;
//             *sheet = Some(new_sheet); // 将新的工作表赋值给 sheet

//             // 写入表头到新工作表
//             for (j, value) in headers.iter().enumerate() {
//                 sheet.as_mut().unwrap().write_string(0, j as u16, value, None)?;
//             }
//             *row_count = 1; // Reset row count after writing header
//         }

//         // 写入数据行
//         for (j, value) in record.iter().enumerate() {
//             sheet.as_mut().unwrap().write_string(*row_count as u32, j as u16, value, None)?;
//         }
//         *row_count += 1;
//     }

//     chunk.clear();
//     Ok(())
// }




// fn process_chunk<'a>(
//     chunk: &Vec<csv::StringRecord>,
//     sheet: &mut Worksheet<'a>,
//     sheet_number: &mut usize,
//     row_count: &mut usize,
//     headers: &csv::StringRecord,
//     rows_per_sheet: usize,
//     workbook: &'a mut Workbook<'a>, // 添加 mut 关键字
// ) -> Result<(), Box<dyn Error>> {
//     for record in chunk {
//         if *row_count >= rows_per_sheet + 1 { // +1 for header
//             // Create new sheet
//             *sheet_number += 1;
//             *sheet = workbook.add_worksheet(Some(&format!("Sheet{}", sheet_number)))?;

//             // Write headers
//             for (j, value) in headers.iter().enumerate() {
//                 sheet.write_string(0, j as u16, value, None)?;
//             }
//             *row_count = 1; // Reset row count after writing header
//         }

//         // Write data row
//         for (j, value) in record.iter().enumerate() {
//             sheet.write_string(*row_count as u32, j as u16, value, None)?;
//         }
//         *row_count += 1;
//     }
//     Ok(())
// }

// 计算CSV文件的总行数（不包括表头）
fn count_csv_rows(file_path: &str) -> Result<usize, Box<dyn Error>> {
    let file = File::open(file_path)?;
    let reader = BufReader::new(file);
    let total_rows = reader.lines().count().saturating_sub(1);  // 使用 saturating_sub 以避免下溢
    Ok(total_rows)
}

// 读取配置文件并解析两个值
fn read_config(file_path: &str) -> Result<(usize, usize), Box<dyn Error>> {
    let file = File::open(file_path)?;
    let reader = BufReader::new(file);

    let mut rows_per_sheet = None;
    let mut chunk_size = None;

    for line in reader.lines() {
        let line = line?;
        if let Some((key, value)) = line.split_once('=') {
            match key.trim() {
                "ROWS_PER_SHEET" => rows_per_sheet = Some(value.trim().parse()?),
                "CHUNK_SIZE" => chunk_size = Some(value.trim().parse()?),
                _ => (),
            }
        }
    }

    let rows_per_sheet = rows_per_sheet.ok_or("ROWS_PER_SHEET 未在配置文件中定义")?;
    let chunk_size = chunk_size.ok_or("CHUNK_SIZE 未在配置文件中定义")?;
    Ok((rows_per_sheet, chunk_size))
}