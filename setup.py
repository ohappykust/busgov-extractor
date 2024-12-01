from cx_Freeze import setup, Executable

setup(
    name="busgov_extractor",
    version="1.0",
    description="Extract data from bus.gov.ru and save it to a XLSX file.",
    executables=[Executable("main.py")]
)
