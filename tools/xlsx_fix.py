from pathlib import Path
from typing import Union


def fix_excel_file_error(path: Union[Path, str]) -> Union[Path, None]:
    import os
    import shutil
    import traceback
    from zipfile import ZipFile

    try:
        file_path = Path(path)
        tmp_folder = file_path.parent.joinpath('__temp__')
        with ZipFile(file_path.__str__()) as excel_container:
            excel_container.extractall(tmp_folder)
            excel_container.close()
        wrong_file_path = os.path.join(tmp_folder.__str__(), 'xl', 'SharedStrings.xml')
        correct_file_path = os.path.join(tmp_folder.__str__(), 'xl', 'sharedStrings.xml')
        os.rename(wrong_file_path, correct_file_path)
        file_path.unlink()
        shutil.make_archive(file_path.__str__(), 'zip', tmp_folder)
        os.rename(file_path.__str__() + '.zip', file_path.__str__())
        shutil.rmtree(tmp_folder.__str__(), ignore_errors=True)
    except (Exception,):
        traceback.print_exc()
        return None
    return file_path


def convert(path_: Path):
    if path_.suffix == '.xls':
        new_path_ = Path(f'{path_}x')
        if new_path_.is_file():
            new_path_.unlink()
        import win32com.client as win32
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(path_.__str__())
        wb.SaveAs(new_path_.__str__(), FileFormat=51)
        wb.Close()
        excel.Application.Quit()
        path_.unlink()
        return new_path_
    else:
        return path_
