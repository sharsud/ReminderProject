from setuptools import setup, find_packages

setup(
    name="learn_with_psudo_reminder",
    version="1.0.0",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "Pillow", "Pandas"
    ],
    entry_points={
        "console_scripts": [
            "learnpsudo-reminder=learn_with_psudo_reminder.main:run_app",
        ]
    },
)


# To Create a build
# D:\Study\Python\ReminderProject> pip install
#  pip install .
# from cmd:
# C:\Users\Sudha\AppData\Local\Programs\Python\Python312\python.exe D:\Study\Python\ReminderProject\LWS_reminder\mainfile.py
