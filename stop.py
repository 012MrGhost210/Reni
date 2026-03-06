import os

# Путь к stop.txt (должен совпадать с основным скриптом)
STOP_FILE = r"D:\мои_письма\stop.txt"

def remove_stop_file():
    """Удаляет файл stop.txt"""
    
    if os.path.exists(STOP_FILE):
        try:
            os.remove(STOP_FILE)
            print(f"✅ Файл {STOP_FILE} успешно удален")
            print("Теперь можно снова запустить отправку писем")
        except Exception as e:
            print(f"❌ Ошибка при удалении: {e}")
    else:
        print(f"❌ Файл {STOP_FILE} не найден")

if __name__ == "__main__":
    remove_stop_file()
pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org win32 -vvv
Using pip 21.2.3 from C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\lib\site-packages\pip (python 3.10)
Non-user install because site-packages writeable
Created temporary directory: C:\Users\Ilya.Matveev2\AppData\Local\Temp\pip-ephem-wheel-cache-4vvdi29a
Created temporary directory: C:\Users\Ilya.Matveev2\AppData\Local\Temp\pip-req-tracker-1owp41g0
Initialized build tracking at C:\Users\Ilya.Matveev2\AppData\Local\Temp\pip-req-tracker-1owp41g0
Created build tracker: C:\Users\Ilya.Matveev2\AppData\Local\Temp\pip-req-tracker-1owp41g0
Entered build tracker: C:\Users\Ilya.Matveev2\AppData\Local\Temp\pip-req-tracker-1owp41g0
Created temporary directory: C:\Users\Ilya.Matveev2\AppData\Local\Temp\pip-install-fs3vr0xo
1 location(s) to search for versions of win32:
* https://pypi.org/simple/win32/
Fetching project page and analyzing links: https://pypi.org/simple/win32/
Getting page https://pypi.org/simple/win32/
Found index url https://pypi.org/simple
Looking up "https://pypi.org/simple/win32/" in the cache
Request header has "max_age" as 0, cache bypassed
Starting new HTTPS connection (1): pypi.org:443
https://pypi.org:443 "GET /simple/win32/ HTTP/1.1" 404 13
Status code 404 not in (200, 203, 300, 301)
Could not fetch URL https://pypi.org/simple/win32/: 404 Client Error: Not Found for url: https://pypi.org/simple/win32/ - skipping
Skipping link: not a file: https://pypi.org/simple/win32/
Given no hashes to check 0 links for project 'win32': discarding no candidates
ERROR: Could not find a version that satisfies the requirement win32 (from versions: none)
ERROR: No matching distribution found for win32
Exception information:
Traceback (most recent call last):
  File "C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\lib\site-packages\pip\_vendor\resolvelib\resolvers.py", line 341, in resolve
    self._add_to_criteria(self.state.criteria, r, parent=None)
  File "C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\lib\site-packages\pip\_vendor\resolvelib\resolvers.py", line 173, in _add_to_criteria
    raise RequirementsConflicted(criterion)
pip._vendor.resolvelib.resolvers.RequirementsConflicted: Requirements conflict: SpecifierRequirement('win32')

During handling of the above exception, another exception occurred:

Traceback (most recent call last):
  File "C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\lib\site-packages\pip\_internal\resolution\resolvelib\resolver.py", line 94, in resolve
    result = self._result = resolver.resolve(
  File "C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\lib\site-packages\pip\_vendor\resolvelib\resolvers.py", line 472, in resolve
    state = resolution.resolve(requirements, max_rounds=max_rounds)
  File "C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\lib\site-packages\pip\_vendor\resolvelib\resolvers.py", line 343, in resolve
    raise ResolutionImpossible(e.criterion.information)
pip._vendor.resolvelib.resolvers.ResolutionImpossible: [RequirementInformation(requirement=SpecifierRequirement('win32'), parent=None)]

The above exception was the direct cause of the following exception:

Traceback (most recent call last):
  File "C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\lib\site-packages\pip\_internal\cli\base_command.py", line 173, in _main
    status = self.run(options, args)
  File "C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\lib\site-packages\pip\_internal\cli\req_command.py", line 203, in wrapper
    return func(self, options, args)
  File "C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\lib\site-packages\pip\_internal\commands\install.py", line 315, in run
    requirement_set = resolver.resolve(
  File "C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\lib\site-packages\pip\_internal\resolution\resolvelib\resolver.py", line 103, in resolve
    raise error from e
pip._internal.exceptions.DistributionNotFound: No matching distribution found for win32
WARNING: You are using pip version 21.2.3; however, version 26.0.1 is available.
You should consider upgrading via the 'C:\Users\Ilya.Matveev2\AppData\Local\Programs\Python\Python310\python.exe -m pip install --upgrade pip' command.
Removed build tracker: 'C:\\Users\\Ilya.Matveev2\\AppData\\Local\\Temp\\pip-req-tracker-1owp41g0'
