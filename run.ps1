# Roda a automação sem instalar o pacote (não gera .egg-info).
$env:PYTHONPATH = "src"
python -m sefaz_ce.main @args
