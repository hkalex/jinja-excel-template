from pathlib import Path
import shutil
import os

dir_path = os.path.dirname(os.path.realpath(__file__))

def pytest_sessionstart():
    """
    Called after the Session object has been created and
    before performing collection and entering the run test loop.
    """
    tempfolder = dir_path + '/temp'
    print("tempfolder=" + tempfolder)
    if os.path.exists(tempfolder) and os.path.isdir(tempfolder):
        shutil.rmtree(tempfolder)
    
    os.mkdir(tempfolder)



