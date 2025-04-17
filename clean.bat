C:
cd\HS\Graph Email Processor v1.1
for /d /r . %%d in (bin,obj,de,es,ja,ru,Express) do @if exist "%%d" rd /s/q "%%d"
