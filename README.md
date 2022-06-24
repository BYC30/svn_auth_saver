## 目的
主要是为了解决svn权限管理不方便容易出错的问题  

## 思路
svn 权限文件本身是有纯文本格式来描述的, 直接编辑的时候容易出错, 并且权限划分不直观  
考虑使用Excel来承载账号信息以及权限信息, 策划可以方便的修改 Excel 表格  
然后通过工具将 Excel 表格的数据转换成 svn 的权限配置格式  
辅助以 Excel vba 在表格保存后调用转换工具  
加上 svn 服务器收到更新后, 更新本地 svn 目录, 实现该目录内的权限自动更新  
最终达到策划修改Excel表格内的权限配置, 保存, 提交即可更新 svn 权限的效果

## 工具命令行参数
工具要求输入两个参数
* input excel文件所在路径
* pwd 导出的svn密码文件名称

## Excel VBA
通过Excel VBA 来传递命令行参数  
```vba
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Shell (ThisWorkbook.Path & "/svn_auth_saver.exe " & Application.ActiveWorkbook.FullName & " svnpasswd")
End Sub
```

## svn hook
通过 svn 的 post-commit 脚本来实现用户提交后更新权限  

```shell
export LANG=en_US.UTF-8 # 字符集，与服务器一致，可执行locale命令查看
svn update /path/to/conf >> /path/to/hooklog # 更新指定目录的svn, 需要将需要管理的svn版本的配置文件指向该目录  
exit 0
```

## 编译
* 安装 rust 环境
* cargo build --release