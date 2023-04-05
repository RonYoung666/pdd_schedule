use chrono::prelude::*;
use chrono::Local;
use chrono::NaiveDate;
use proconio::input;
use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, Workbook, XlsxError};
use sscanf::sscanf;
use std::io;

/* 班种：白班、晚班、中班、休息 */
#[derive(Copy, Clone, Debug, PartialEq)]
#[allow(dead_code)]
enum Shift {
    Unset,
    Day,
    Eve,
    Mid,
    Rest,
}

/* 一个员工的班表 */
#[derive(Clone, Debug)]
#[allow(dead_code)]
struct Employee {
    name: String,
    shift: [Shift; 31],
    day_shift_num: usize,
    eve_shift_num: usize,
    mid_shift_num: usize,
    rest_shift_num: usize,
}

/* 一天的各班次人数 */
#[derive(Copy, Clone)]
#[allow(dead_code)]
struct ScheduleCountOneDay {
    day: usize,
    eve: usize,
    mid: usize,
    rest: usize,
}

/* 下个月的 1 号减这个月的 1 号得到这个月的天数 */
fn get_days_from_month(year: i32, month: u32) -> i64 {
    let next_month = NaiveDate::from_ymd_opt(
        match month {
            12 => year + 1,
            _ => year,
        },
        match month {
            12 => 1,
            _ => month + 1,
        },
        1,
    );
    let this_month = NaiveDate::from_ymd_opt(year, month, 1);

    next_month
        .unwrap()
        .signed_duration_since(this_month.unwrap())
        .num_days()
}

fn gen_xlsx(
    year: i32,
    month: u32,
    day_sum: usize,
    employee_num: usize,
    employee: &Vec<Employee>,
    schedule_count: &[ScheduleCountOneDay; 31],
) -> Result<(), XlsxError> {
    /* 新建 Excel 对象 */
    let mut workbook = Workbook::new();

    /* 创建一些格式 */
    let title_format = Format::new()
        .set_bold()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center);
    let normal_format = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center);
    let rest_format = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_font_color(rust_xlsxwriter::XlsxColor::Red);

    /* 新建一个 Sheet */
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("拼多多")?;

    /* 把日期单元格放窄 */
    for col in 1..(day_sum as u16 + 1) {
        worksheet.set_column_width(col, 2.5)?;
    }

    /* X 月份拼多多班表 */
    worksheet.merge_range(
        0,
        1,
        0,
        day_sum as u16,
        &format!("{month}月份拼多多班表")[..],
        &title_format,
    )?;

    /* 日期 */
    worksheet.write_string_with_format(1, 0, "日期", &normal_format)?;
    for col in 1..(day_sum as u16 + 1) {
        worksheet.write_string_with_format(1, col, &format!("{col}")[..], &title_format)?;
    }

    /* 星期 */
    for i in 0..day_sum {
        let weekday = Utc
            .with_ymd_and_hms(year, month, i as u32 + 1, 0, 0, 0)
            .unwrap()
            .weekday();
        let cn_weekday = ["一", "二", "三", "四", "五", "六", "日"];
        worksheet.write_string_with_format(
            2,
            i as u16 + 1,
            cn_weekday[weekday as usize],
            &normal_format,
        )?;
    }

    /* 白班、晚班、中班、休假、年假、余假 */
    let shift_title_format = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter);
    let shift_title = ["白班", "晚班", "中班", "休假", "年假", "余假"];
    for i in 0..6 {
        let col = day_sum as u16 + i as u16 + 1;
        worksheet.set_column_width(col, 4)?;
        worksheet.merge_range(1, col, 2, col, shift_title[i], &shift_title_format)?;
    }

    /* 各个员工的班表 */
    for i in 0..(employee.len()) {
        worksheet.write_string_with_format(
            i as u32 + 3,
            0,
            &format!("{}", employee[i].name)[..],
            &normal_format,
        )?;
        for j in 0..day_sum {
            let shift = ["", "白", "晚", "中", "休"];
            let tmp_fmt = [
                normal_format.clone(),
                normal_format.clone(),
                normal_format.clone(),
                normal_format.clone(),
                rest_format.clone(),
            ];
            worksheet.write_string_with_format(
                i as u32 + 3,
                j as u16 + 1,
                shift[employee[i].shift[j] as usize],
                &tmp_fmt[employee[i].shift[j] as usize],
            )?;
        }
        let num = [
            employee[i].day_shift_num,
            employee[i].eve_shift_num,
            employee[i].mid_shift_num,
            employee[i].rest_shift_num,
            0,
            0,
        ];
        for j in 0..6 {
            worksheet.write_string_with_format(
                i as u32 + 3,
                day_sum as u16 + j as u16 + 1,
                &format!("{}", num[j]),
                &normal_format,
            )?;
        }
    }

    /* 最下面的统计 */
    let shift_title = ["白班", "晚班", "中班", "休"];
    for i in 0..4 {
        worksheet.write_string_with_format(
            employee_num as u32 + i as u32 + 3,
            0,
            shift_title[i],
            &title_format,
        )?;
    }
    for i in 0..day_sum {
        worksheet.write_string_with_format(
            employee_num as u32 + 3,
            i as u16 + 1,
            &format!("{}", schedule_count[i].day),
            &normal_format,
        )?;
        worksheet.write_string_with_format(
            employee_num as u32 + 4,
            i as u16 + 1,
            &format!("{}", schedule_count[i].eve),
            &normal_format,
        )?;
        worksheet.write_string_with_format(
            employee_num as u32 + 5,
            i as u16 + 1,
            &format!("{}", schedule_count[i].mid),
            &normal_format,
        )?;
        worksheet.write_string_with_format(
            employee_num as u32 + 6,
            i as u16 + 1,
            &format!("{}", schedule_count[i].rest),
            &normal_format,
        )?;
    }

    /* 保存 Excel 文件 */
    workbook.save(format!(
        "{year}年{month}月份拼多多班表_{}.xlsx",
        Local::now().format("%Y-%m-%d_%H%M%S").to_string()
    ))?;

    Ok(())
}

/*
 * 获取最大的连续休息天数
 * 返回值：(开始索引, 天数)
 */
fn get_max_rest(day_sum: usize, employee: &Employee) -> (usize, usize) {
    let mut max_index = 0;
    let mut max_num = 0;

    /* 遍历一个月每一天 */
    for i in 0..day_sum {
        /* 第一个休 */
        if employee.shift[i] == Shift::Rest {
            /* 向后累加 */
            for j in i + 1..day_sum {
                if employee.shift[j] != Shift::Rest {
                    if j - i > max_num {
                        max_index = i;
                        max_num = j - i;
                    }
                    break;
                }
            }
        }
    }

    return (max_index, max_num);
}

fn main() {
    /* 输入年月 */
    println!("请输入年月，如 \"2023-03\"");
    let mut ym = String::new();
    io::stdin().read_line(&mut ym).expect("获取年月失败！");
    let ym = ym.trim();
    let (year, month) = match sscanf!(ym, "{}-{}", i32, u32) {
        Ok((year, month)) => (year, month),
        Err(_) => {
            // println!("获取输入失败！Error = [{:?}]", msg);
            println!("请按照格式 \"YYYY-MM\" 输入!");
            return;
        }
    };
    if year < 0 || month <= 0 || month > 12 {
        println!("请输入正确的年月!");
        return;
    }
    println!("将生成 {year} 年 {month} 月的团队排期表");

    /* 获取一个月的天数 */
    let day_sum = get_days_from_month(year, month) as usize;
    // println!("day_sum = {}", day_sum);

    /* 输入员工个数 */
    println!("请输入员工个数");
    input! {
        employee_num: usize,
    }
    // println!("employee_num = {employee_num}");

    /* 申请员工信息数组 */
    let mut employee = vec![
        Employee {
            name: String::new(),
            shift: [Shift::Unset; 31],
            day_shift_num: 0,
            eve_shift_num: 0,
            mid_shift_num: 0,
            rest_shift_num: 0
        };
        employee_num
    ];

    /* 输入多个员工的休假日期 */
    println!("请输入员工姓名、休假天数和休假日期，如 \"刘琼 5 5 12 18 19 26\"，一行一个员工");
    for i in 0..employee.len() {
        println!("请输入员工 {} 的数据：", i + 1);

        input! {
            name: String,
            rest_shift_num: usize,
            rest_days: [usize; rest_shift_num],
        }

        employee[i].name = name;
        employee[i].rest_shift_num = rest_shift_num;
        for rest_day in rest_days {
            employee[i].shift[rest_day - 1] = Shift::Rest;
        }
    }

    /* 计算每个人的各种班总数 */
    for e in &mut employee {
        /* 刘琼的特殊处理 */
        if e.name.eq("刘琼") {
            e.eve_shift_num = 4;
            e.day_shift_num = day_sum - e.rest_shift_num - e.eve_shift_num;
            continue;
        }

        e.day_shift_num = (day_sum - e.rest_shift_num) / 2;
        e.eve_shift_num = day_sum - e.rest_shift_num - e.day_shift_num;
    }

    /* 排班 */
    for e in &mut employee {
        let (start_index, con_num) = get_max_rest(day_sum, e);

        let mut eve_left = e.eve_shift_num; /* 晚班剩余 */
        let mut day_left = e.day_shift_num; /* 白班剩余 */

        /* 从休息之后开始排，先晚再白 */
        let mut cur_shift = Shift::Eve;
        let mut i = start_index + con_num;
        while i < day_sum {
            /* 班次剩余为 0，换班 */
            if cur_shift == Shift::Eve && eve_left == 0 {
                cur_shift = Shift::Day;
            }
            if cur_shift == Shift::Day && day_left == 0 {
                cur_shift = Shift::Eve;
            }

            /* 遇到休息进行转换 */
            if e.shift[i] == Shift::Rest {
                /* 换班 */
                if cur_shift == Shift::Eve && day_left > 0 {
                    cur_shift = Shift::Day;
                } else {
                    cur_shift = Shift::Eve;
                }

                /* 跳过休息的天数 */
                for j in i + 1..day_sum {
                    if e.shift[j] != Shift::Rest {
                        i = j - 1;
                        break;
                    }
                }

                i += 1;
                continue;
            }

            /* 排班 */
            if cur_shift == Shift::Eve {
                e.shift[i] = Shift::Eve;
                eve_left -= 1;
            } else {
                e.shift[i] = Shift::Day;
                day_left -= 1;
            }

            i += 1;
        }

        /* 从休息之前开始排，先白再晚 */
        let mut cur_shift = Shift::Day;
        let mut i = start_index.checked_sub(1).unwrap_or(0);
        loop {
            /* 班次剩余为 0，换班 */
            if cur_shift == Shift::Eve && eve_left == 0 {
                cur_shift = Shift::Day;
            }
            if cur_shift == Shift::Day && day_left == 0 {
                cur_shift = Shift::Eve;
            }

            /* 遇到休息进行转换 */
            if e.shift[i] == Shift::Rest {
                /* 换班 */
                if cur_shift == Shift::Eve && day_left > 0 {
                    cur_shift = Shift::Day;
                } else {
                    cur_shift = Shift::Eve;
                }

                /* 跳过休息的天数 */
                for j in (0..i).rev() {
                    if e.shift[j] != Shift::Rest {
                        i = j + 1;
                        break;
                    }
                }

                if i == 0 {
                    break;
                }
                if i > 0 {
                    i -= 1;
                }
                continue;
            }

            /* 排班 */
            if cur_shift == Shift::Eve {
                e.shift[i] = Shift::Eve;
                eve_left = eve_left.checked_sub(1).unwrap_or(0);
            } else {
                e.shift[i] = Shift::Day;
                day_left = day_left.checked_sub(1).unwrap_or(0);
            }

            if i == 0 {
                break;
            }
            if i > 0 {
                i -= 1;
            }
        }

        /* 如果有晚转白，白换成中 */
        for i in 0..day_sum - 1 {
            if e.shift[i] == Shift::Eve && e.shift[i + 1] == Shift::Day {
                e.shift[i + 1] = Shift::Mid;
                e.day_shift_num -= 1;
                e.mid_shift_num += 1;
            }
        }
    }

    /* 每天的班次统计 */
    let mut schedule_count = [ScheduleCountOneDay {
        day: 0,
        eve: 0,
        mid: 0,
        rest: 0,
    }; 31];
    for i in 0..day_sum {
        for e in &employee {
            match e.shift[i] {
                Shift::Day => schedule_count[i].day += 1,
                Shift::Eve => schedule_count[i].eve += 1,
                Shift::Mid => schedule_count[i].mid += 1,
                Shift::Rest => schedule_count[i].rest += 1,
                _ => (),
            }
        }
    }

    /* 打印排班信息到屏幕 */
    print!("日期  ");
    for i in 0..day_sum {
        print!(" {:2}", i + 1);
    }
    print!(" 白班 晚班 中班 休假\n");
    for e in &employee {
        print!("{}", e.name);
        if e.name.len() == 6 {
            print!("  ");
        }

        for i in 0..day_sum {
            match e.shift[i] {
                Shift::Day => print!(" 白"),
                Shift::Mid => print!(" 中"),
                Shift::Eve => print!(" 晚"),
                Shift::Rest => print!(" \x1B[31m休\x1B[0m"),
                _ => print!("   "),
            }
        }

        print!("{:5}", e.day_shift_num);
        print!("{:5}", e.eve_shift_num);
        print!("{:5}", e.mid_shift_num);
        print!("{:5}", e.rest_shift_num);

        println!("");
    }
    print!("白班  ");
    for i in 0..day_sum {
        print!("{:3}", schedule_count[i].day);
    }
    println!("");
    print!("晚班  ");
    for i in 0..day_sum {
        print!("{:3}", schedule_count[i].eve);
    }
    println!("");
    print!("中班  ");
    for i in 0..day_sum {
        print!("{:3}", schedule_count[i].mid);
    }
    println!("");
    print!("休    ");
    for i in 0..day_sum {
        print!("{:3}", schedule_count[i].rest);
    }
    println!("");

    match gen_xlsx(
        year,
        month,
        day_sum,
        employee_num,
        &employee,
        &schedule_count,
    ) {
        Err(e) => println!("{:?}", e),
        _ => (),
    }
}
