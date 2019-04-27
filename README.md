## 简介

将sql语句查询结果导出至excel文件

## 安装
```
cd $GOPATH

go get -v github.com/wanglilong2013/mysql2Xlsx
```
## 编译
```
cd $GOPATH/src/github.com/wanglilong/2013/mysql2Xlsx

go build

```

## 运行
```
./mysql2Xlsx  -h localhost -P 3306  -d gogs -u root -t ./user2.xlsx
```
