# Development Notes

## 推荐位置

项目推荐放在：

```text
/Users/neoshi/Developer/Invoicely
```

`dist/` 是成品输出目录，不建议提交到 git。需要发布或复制给其他电脑时，重新运行 `./scripts/build_app.sh` 生成即可。

## App 本地数据

App 的设置和汇率缓存会保存在：

```text
~/Library/Application Support/Invoicely
```

这是正常 macOS App 数据。删除该目录会重置设置和汇率缓存，不影响源码项目。
