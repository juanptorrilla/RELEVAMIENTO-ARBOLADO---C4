from pathlib import Path
import re

android = Path('android')
main_files = list(android.rglob('MainActivity.java'))
if len(main_files) != 1:
    raise SystemExit(f'Se esperaba un MainActivity.java y se encontraron {len(main_files)}')
main = main_files[0]
source = Path('native/DownloadsPlugin.java')
target = main.parent / 'DownloadsPlugin.java'
target.write_text(source.read_text(encoding='utf-8'), encoding='utf-8')
text = main.read_text(encoding='utf-8')
if 'DownloadsPlugin.class' not in text:
    text = text.replace('import com.getcapacitor.BridgeActivity;', 'import com.getcapacitor.BridgeActivity;\nimport android.os.Bundle;')
    pattern = r'public class MainActivity extends BridgeActivity\s*\{'
    replacement = '''public class MainActivity extends BridgeActivity {
    @Override
    public void onCreate(Bundle savedInstanceState) {
        registerPlugin(DownloadsPlugin.class);
        super.onCreate(savedInstanceState);
    }
'''
    text, count = re.subn(pattern, replacement, text, count=1)
    if count != 1:
        raise SystemExit('No se pudo registrar DownloadsPlugin en MainActivity.java')
    main.write_text(text, encoding='utf-8')
print(f'Plugin de Descargas instalado en {target} y registrado en {main}')
