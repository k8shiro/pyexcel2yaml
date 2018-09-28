# pyexcel2yaml
SHIFT ware( https://github.com/SHIFT-ware/shift_ware )のExcel2YAMLをinventoryファイル、host_varsファイルにパースするためのpythonスクリプト

# 使い方

pye2y/src/にExcel2YAMLを配置し以下を実行。
※ docker, docker-composeが必要です。

```
docker-compose up -d
docker-compose exec pye2y python /src/pyexcel2yaml.py
```

pye2y/src/export_files/にinventoryファイル、host_varsファイルが出力されます。
または、pye2y/src以下が本体なので、pye2y/src/requirements.txtをpip install後pye2y/src/pyexcel2yaml.pyを実行することでも同様に実行可能です。
