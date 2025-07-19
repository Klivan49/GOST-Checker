import csv

def generate_report(issues, report_path):
    """
    Генерирует CSV-отчёт со списком найденных и исправленных проблем.
    """
    with open(report_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['№', 'Описание проблемы'])
        for idx, issue in enumerate(issues, start=1):
            writer.writerow([idx, issue])
            
        