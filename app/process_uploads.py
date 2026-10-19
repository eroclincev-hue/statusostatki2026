#!/usr/bin/env python3
"""Новые выгрузки ВМС из папки uploads/<ФФ>/ -> база остатков ФФ (ff_snapshots.json).

Время выгрузки = время, когда файл добавили в репозиторий (коммит). После обработки файл
переносится в uploads/processed/. Запускается из GitHub Actions перед обновлением данных.
"""
import os, sys, subprocess, shutil, datetime
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import ffreal

FOLDERS = {'ff1': 'ФФ №1', 'ff2b': 'ФФ2Б', 'flakony': 'Флаконы', 'onyx': 'ONYX FF', 'pitsunov': ffreal.PITSUNOV}


def commit_time(path):
    try:
        t = subprocess.check_output(['git', 'log', '-1', '--format=%cI', '--', path], cwd=ROOT, text=True).strip()
        if t: return datetime.datetime.fromisoformat(t).astimezone(datetime.timezone.utc).isoformat(timespec='seconds')
    except Exception:
        pass
    return datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='seconds')


def main():
    done = 0
    for folder, ff in FOLDERS.items():
        d = os.path.join(ROOT, 'uploads', folder)
        if not os.path.isdir(d): continue
        ext = ('.xlsx', '.xls') if folder == 'pitsunov' else ('.xlsx',)
        files = sorted([f for f in os.listdir(d) if f.lower().endswith(ext)], key=lambda f: commit_time(os.path.join(d, f)))
        if folder == 'pitsunov' and any(f.lower().endswith('.xls') for f in files):
            try:
                import xlrd  # noqa: F401
            except ImportError:
                subprocess.run([sys.executable, '-m', 'pip', 'install', 'xlrd', '--break-system-packages', '-q'], check=False)
        for f in files:  # если файлов несколько, последней станет самая свежая выгрузка
            p = os.path.join(d, f)
            when = commit_time(p)
            print(f'{ff}: {f} от {when}')
            if folder == 'pitsunov': ffreal.load_pitsunov(p, when)
            else: ffreal.load(ff, p, when)
            dst = os.path.join(ROOT, 'uploads', 'processed', f'{folder}_{when[:16].replace(":", "-")}_{f}')
            shutil.move(p, dst)
            done += 1
    print('обработано выгрузок:', done)


if __name__ == '__main__':
    main()
