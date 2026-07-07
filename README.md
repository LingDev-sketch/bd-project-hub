# BD开改关进度管理与自动校验系统

这是一个轻量版 Streamlit 小程序，用来替代每月邮件收取 9 个区域 BD 的新开、整改、关店滚动表。

## 功能

- BD前台提交新开、整改、关店项目进度
- 后台导入系统店铺资料表
- 后台可选导入历史 BD Tracking 三张提交表
- 自动匹配系统店铺资料
- 自动校验进度、时间、系统状态、Region一致性
- 自动生成高风险/需确认异常清单
- 一键导出全国项目库、异常清单、区域进度汇总

## 运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

后台默认密码：

```text
bd-admin
```

可在 `app.py` 顶部修改 `ADMIN_PASSWORD`。

## 推荐使用方式

第一阶段：

1. 每月上传最新系统店铺资料。
2. 导入历史 BD Tracking。
3. 让 BD 从前台提交后续变更。
4. 后台只追异常清单，不再逐行人工检查。

第二阶段：

如果要多人长期使用，建议把数据存储从 SQLite 升级到飞书多维表格、Google Sheet 或正式数据库。
