# 通过subject查找并删除邮件。

### 准备
1. 管理员运行powershell，设置执行策略
`Set-ExecutionPolicy RemoteSigned` 

2. 安装模块
`Install-Module -Name ExchangeOnlineManagement` 

### Usage
1. 首先需要获取到凭据
` .\DeleteMailsBySome.ps1 -object "test_delete_mail_by_powershell" -searchName "jiaoyang_on_new_mail36" ` 

#### parameters
1. object 要搜索邮件的主题
2. searchName 指定本次搜索的名称，任意即可

### TODO
- 新增搜索条件，添加使用发件人进行搜索;
- 凭据获取的健壮性