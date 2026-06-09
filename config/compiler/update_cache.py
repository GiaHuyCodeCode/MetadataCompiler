import json

cache_file = "config/compiler/ai_cache.json"

try:
    with open(cache_file, "r") as f:
        cache = json.load(f)
except FileNotFoundError:
    cache = {}

rows_elec = [{"項目名": "electricity_contact_name", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "電気業者名", "処理": "そのまま出力", "条件 / 項目マッピング": "連絡不要'が含まれていない場合", "開発用チェック": "True"}]
rows_water = [{"項目名": "drink_water_contact_name", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "上水道業者名", "処理": "そのまま出力", "条件 / 項目マッピング": "連絡不要'が含まれていない場合", "開発用チェック": "True"}]
rows_gas = [{"項目名": "gas_contact_name", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "ガス業者名", "処理": "そのまま出力", "条件 / 項目マッピング": "連絡不要'が含まれていないまたは\n\"田中総合燃料\"が含まれていない場合", "開発用チェック": "True"}]
rows_drain = [{"項目名": "drainage_contact_name", "該当ファイル名": "GMO 入居状況一覧.csv", "該当項目名_1": "排水業者名", "処理": "そのまま出力", "条件 / 項目マッピング": "連絡不要'が含まれていない場合または\n\"汲み取り業者\"が含まれていない場合", "開発用チェック": "True"}]

vba_elec = """SQL = ""
SQL = SQL & "UPDATE Contract AS T "
SQL = SQL & "INNER JOIN [GMO 入居状況一覧] AS T1 ON T.ID = ('K_' & T1.ID) "
SQL = SQL & "SET T.electricity_contact_name = T1.[電気業者名] "
SQL = SQL & "WHERE T1.[電気業者名] NOT LIKE '%連絡不要%'; "
db.Execute SQL
log_write "set_contract:契約_K → 特殊処理: electricity_contact_name\""""

vba_water = """SQL = ""
SQL = SQL & "UPDATE Contract AS T "
SQL = SQL & "INNER JOIN [GMO 入居状況一覧] AS T1 ON T.ID = ('K_' & T1.ID) "
SQL = SQL & "SET T.drink_water_contact_name = T1.[上水道業者名] "
SQL = SQL & "WHERE T1.[上水道業者名] NOT LIKE '%連絡不要%'; "
db.Execute SQL
log_write "set_contract:契約_K → 特殊処理: drink_water_contact_name\""""

vba_gas = """SQL = ""
SQL = SQL & "UPDATE Contract AS T "
SQL = SQL & "INNER JOIN [GMO 入居状況一覧] AS T1 ON T.ID = ('K_' & T1.ID) "
SQL = SQL & "SET T.gas_contact_name = T1.[ガス業者名] "
SQL = SQL & "WHERE T1.[ガス業者名] NOT LIKE '%連絡不要%' OR T1.[ガス業者名] NOT LIKE '%田中総合燃料%'; "
db.Execute SQL
log_write "set_contract:契約_K → 特殊処理: gas_contact_name\""""

vba_drain = """SQL = ""
SQL = SQL & "UPDATE Contract AS T "
SQL = SQL & "INNER JOIN [GMO 入居状況一覧] AS T1 ON T.ID = ('K_' & T1.ID) "
SQL = SQL & "SET T.drainage_contact_name = T1.[排水業者名] "
SQL = SQL & "WHERE T1.[排水業者名] NOT LIKE '%連絡不要%' OR T1.[排水業者名] NOT LIKE '%汲み取り業者%'; "
db.Execute SQL
log_write "set_contract:契約_K → 特殊処理: drainage_contact_name\""""

cache[json.dumps(rows_elec, sort_keys=True)] = vba_elec
cache[json.dumps(rows_water, sort_keys=True)] = vba_water
cache[json.dumps(rows_gas, sort_keys=True)] = vba_gas
cache[json.dumps(rows_drain, sort_keys=True)] = vba_drain

with open(cache_file, "w") as f:
    json.dump(cache, f, indent=2, ensure_ascii=False)

print("Updated cache!")
