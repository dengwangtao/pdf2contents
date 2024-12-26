# 东华大学学报
# 提取pdf生成目录(中英文标题\作者)

## 处理PDF， 获得目录

```shell
mkdir build

cd build

pyinstaller -F -n 'preocessPDF.exe' -i ../logo.png ../main2.py
```