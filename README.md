
## API
### 修改表格指定位置内容
```
POST /api/ChangeTextInCell/ChangeText

{
    "docUrl": "", //文件地址
    "rowIndex": 3, //第几行
    "columnIndex": 3, //第几列
    "content": "  测试123456" //修改的内容
}

```

### word 转 pdf
```
POST /api/WordToPdf/WToP

{
    "sourcedocx":"", //word 文件路劲
    "targetpdf":"" //保存pdf文件路劲
}
```

### word 转 pdf
```
get /api/IOFile/DownloadFile?fn="文件名字"

{
    "sourcedocx":"", //word 文件路劲
    "targetpdf":"" //保存pdf文件路劲
}
```


## 常见问题
### win7 本地 IIS6服务 建立Microsoft.Office.Interop.Word.Application时出现“拒绝访问”错误的解决方法   [参考文档](https://www.cnblogs.com/firstyi/articles/1277307.html)

  - 1 
    - 在IIS管理器 -> 应用程序池 -> 选择对应项目池子右键 -> 高级设置 -> 标识 -> 内置账号设置为 LocalSystem
  
  - 2
    - 1、在命令行中输入：dcomcnfg，会显示出“组件服务”管理器

    - 2、打开“组件服务->计算机->我的电脑->DCOM 配置”，找到“Microsoft Word文档”，单击右键，选择“属性”

    - 3、在“属性”对话框中单击“标识”选项卡，选择“交互式用户””

  - 备注 如上述两项未能解决问题, 在尝试下面第三项

  - 3
    - 1、在命令行中输入：dcomcnfg，会显示出“组件服务”管理器

    - 2、打开“组件服务->计算机->我的电脑->DCOM 配置”，找到“Microsoft Word文档”，单击右键，选择“属性”

    - 3、在“属性”对话框中单击“安全”选项卡，在“启动和激活权限”处选择“自定义”，再单击右边的“编辑”，在弹出的对话框中添加“ASPNET”（在IIS6中是NETWORD SERVICE）用户，给予“本地启动”和“本地激活”的权限，单击“确定

    - 4、在“属性”对话框中单击“安全”选项卡，在“访问权限”处选择“自定义”，再单击右边的“编辑”，在弹出的对话框中添加“ASPNET”（在IIS6中是NETWORD SERVICE）用户，给予“本地访问”的权限，单击“确定”，关闭“组件服务”管理器。
