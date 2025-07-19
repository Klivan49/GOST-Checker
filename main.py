import sys
import os
import yaml
from docx import Document
from checker.style_checker import check_and_fix_style
from checker.report_generator import generate_report
from checker.utils import process_bibliography

CONFIG_PATH = os.path.join(os.path.dirname(__file__), 'config.yaml')

def load_config(path):
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)

def main(docx_path):
    if not os.path.isfile(docx_path):
        print(f"[!] Файл не найден: {docx_path}")
        sys.exit(1)

    config = load_config(CONFIG_PATH)
    doc = Document(docx_path)

    issues = check_and_fix_style(doc, config)

    process_bibliography(doc, config.get('bibliography', {}))

    base, ext = os.path.splitext(docx_path)
    output_path = f"{base}_gost{ext}"
    doc.save(output_path)
    print(f"[✓] Исправленный документ сохранён: {output_path}")

    if issues:
        report_path = f"{base}_report.csv"
        generate_report(issues, report_path)
        print(f"[!] Обнаружены проблемы. Смотрите отчёт: {report_path}")
    else:
        print("[✓] Проблем не обнаружено.")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Использование: python main.py <путь_к_документу.docx>")
        sys.exit(1)
    main(sys.argv[1])