from datetime import date
from pathlib import Path
from typing import List

username: str = ""
password: str = ""

username_friend_list: List[str] = ["", ""]

path_excel: Path = Path("")
name_file_excel: str = ""

name_sheet: str = ""
name_column: List[str] = ["Date", "Lunch", "Food", "Calories", "Protein", "Fat", "Carbohydrates", "Sodium", "Sugar"]

safe_mode: bool = False

## Catching-up
manual_date_mode: bool = True
manual_date: List[date] = [date(2022, 1, 18), date(2022, 1, 20)]

