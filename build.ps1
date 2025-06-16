$exclude = @("venv", "botPython.zip", "#material", "downloads")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "botPython.zip" -Force