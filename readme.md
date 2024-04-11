# 通过搜索表达式搜索邮件及删除功能

### 准备
1. 管理员运行powershell，设置执行策略
`Set-ExecutionPolicy RemoteSigned` 

2. 安装模块
`Install-Module -Name ExchangeOnlineManagement` 

### Usage
1. 首先需要获取到凭据
` .\DeleteMailsBySome.ps1 -searchExpression "搜索表达式" -searchName "搜索任务名" ` 

#### parameters
1. searchExpression 搜索表达式, eg. `'(Received=2/28/2024) AND (Subject:"RecallTest2")'`
2. searchName 指定本次搜索的名称，任意即可

