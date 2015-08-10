
#Standard Library
import glob

#Local
import cleanup_ccli


if __name__ == '__main__':

    folders = [
        r'C:\Users\avteam\Google Drive\RMPC PowerPoint\PowerPoint Songs',
    ]
    
    for folder in folders:
        for powerpoint in glob.iglob(folder + '/*.ppt*'):
            print powerpoint
        
    
    

