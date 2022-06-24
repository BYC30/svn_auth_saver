use std::{path::PathBuf, fs, collections::{HashMap, HashSet}};

use calamine::{open_workbook_auto, Reader, Sheets};
use structopt::StructOpt;
use anyhow::{Result, anyhow};
use thiserror::Error;
use msgbox::IconType;

#[derive(Error, Debug)]
pub enum CsvError {
    #[error("页签[{0}]未找到")]
    SheetNotFound(String),
    #[error("Excel打开错误")]
    ExcelError(#[from] calamine::XlsxError),
    #[error("用户[{0}]没有定义在passwd中")]
    UserNotExist(String),
    #[error("用户[{0}]分组重复定义")]
    UserDupAuth(String),
    #[error("用户[{0}]密码为空")]
    UserEmptyPwd(String),
    #[error("无效权限[{0}]配置")]
    InvalidAuth(String),
    #[error("用户组[{0}]重复")]
    GroupNameDup(String),
    #[error("未知错误[{0}]")]
    Other(#[from] anyhow::Error),
}

#[derive(Debug, StructOpt)]
#[structopt(name = "excel2csv", about = "export excel to csv")]
struct Opt {
    /// excel file that need to convert
    #[structopt(parse(from_os_str))]
    input: PathBuf,
    pwd: String,
}

fn save_pwd(workbook: &mut Sheets, name: &str, p: &str) -> Result<HashSet<String>> {
    let mut content = vec![];
    content.push("[users]".to_string());
    let mut ret = HashSet::new();
    let range = workbook.worksheet_range(name)
        .ok_or(CsvError::SheetNotFound(name.to_string()))??;
    for row in range.rows() {
        let name = row[0].to_string();
        let passwd = row[1].to_string();
        if name.is_empty() {continue;}
        if passwd.is_empty() {
            return Err(anyhow!(CsvError::UserEmptyPwd(name)));
        }
        content.push(format!("{}={}", name, passwd));
        ret.insert(name.to_string());
    }
    let total = content.join("\r\n");
    println!("save {:?}", p);
    fs::write(p, total)?;
    Ok(ret)
}

fn save_auth(workbook: &mut Sheets, name: &str, p: &str, name_set:&HashSet<String>) -> Result<()> {
    let mut user = vec![];
    user.push("[groups]".to_string());
    let mut auth = vec![];
    let range = workbook.worksheet_range(name)
        .ok_or(CsvError::SheetNotFound(name.to_string()))??;
    let mut row_idx = 0;
    let mut col_map = HashMap::new();
    let mut group_order = Vec::new();
    let mut group_map:HashMap<String, Vec<String>> = HashMap::new();
    let mut group_user = HashSet::new();
    for row in range.rows() {
        row_idx = row_idx + 1;
        match row_idx {
            1 => {
                let mut idx = 0;
                for one in row{
                    idx = idx + 1;
                    if one.is_empty(){continue;}
                    let v = one.to_string();
                    if group_map.contains_key(&v) {
                        return Err(anyhow!(CsvError::GroupNameDup(v)));
                    }
                    group_order.push(v.clone());
                    group_map.insert(v.clone(), vec![]);
                    col_map.insert(idx, v.clone());
                }
            },
            _ => {
                let row_key = row[0].to_string();
                if row_key.is_empty() {continue;}
                if !row_key.starts_with("/") {
                    let mut idx = 0;
                    let user_name = row_key;
                    if !name_set.contains(&user_name) {
                        return Err(anyhow!(CsvError::UserNotExist(user_name)));
                    }
                    if group_user.contains(&user_name) {
                        return Err(anyhow!(CsvError::UserDupAuth(user_name)));
                    }
                    group_user.insert(user_name.clone());
                    for one in row{
                        idx = idx + 1;
                        if idx == 1 {continue;}
                        if one.is_empty(){continue;}
                        let name = col_map.get(&idx);
                        if name.is_none() {continue;}
                        let name = name.unwrap();
                        let  group_member = group_map.get_mut(name);
                        if group_member.is_none() {continue;}
                        let group_member = group_member.unwrap();
                        group_member.push(user_name.clone());
                    }
                }else{
                    let mut idx = 0;
                    for one in row{
                        idx = idx + 1;
                        match idx {
                            1 => {
                                if one.is_empty(){break;}
                                auth.push(format!("[{}]", one.to_string()));
                                auth.push("*=".to_string());
                            },
                            _ => {
                                if one.is_empty(){continue;}
                                let v = one.to_string();
                                if v != "r" && v != "rw" {
                                    return Err(anyhow!(CsvError::InvalidAuth(v)));
                                }
                                let name = col_map.get(&idx);
                                if name.is_none() {continue;}
                                let name = name.unwrap();
                                let one_line = format!("{}={}", name, v);
                                auth.push(one_line);
                            }
                        }
                    }
                    auth.push("\r\n".to_string());
                }
            }
        }
    }
    let mut total = "".to_string();
    for one in group_order {
        let members = group_map.get(&one);
        if members.is_none() {continue;}
        let members = members.unwrap();
        let one_line = format!("{}={}", one.replace("@", ""), members.join(","));
        user.push(one_line)
    }
    user.push("\r\n".to_string());
    let user_txt = user.join("\r\n");
    total.push_str(&user_txt);
    let auth_txt = auth.join("\r\n");
    total.push_str(&auth_txt);
    println!("save {:?}", p);
    fs::write(p, total)?;
    Ok(())
}

fn save(opt:&Opt) -> Result<()> {
    let path = &opt.input;
    let mut workbook= open_workbook_auto(path)?;
    let mut pwd_path = path.clone();
    pwd_path.pop();
    pwd_path.push(opt.pwd.as_str());

    let name_set = save_pwd(&mut workbook, "passwd", pwd_path.to_str().unwrap())?;

    let range = workbook.worksheet_range("config")
        .ok_or(CsvError::SheetNotFound("config".to_string()))??;
    let mut idx = 0;
    for row in range.rows() {
        idx = idx + 1;
        if idx <= 1 {continue;}
        let name = match row[0].get_string() {
            None => continue,
            Some(x) => x,
        };
        let save = match row[1].get_string() {
            None => continue,
            Some(x) => x,
        };
        let mut authz_path = path.clone();
        authz_path.pop();
        authz_path.push(save);
    
        save_auth(&mut workbook, name, authz_path.to_str().unwrap(), &name_set)?;
    }

    
    Ok(())
}

fn main() -> Result<()> {
    let opt = Opt::from_args();

    let f = save(&opt);
    match f {
        Ok(())=>{
            msgbox::create("导出完毕", "导出成功", IconType::Info)?;
        },
        Err(e) => {
            msgbox::create("导出完毕", e.to_string().as_str(), IconType::Info)?;
        }
    }
    Ok(())
}
