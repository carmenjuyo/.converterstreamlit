from datetime import datetime
import sys, os


def writeFile(x=None):
    t = datetime.utcnow()
    print(t)
    if x == None:
        with open('log/log/log/log.txt', 'w') as f: f.write(str(f'{t} - {x}') + "\n")
    else:
        with open('log/log/log/log.txt', 'a') as f: f.write(str(f'{t} - {x}') + "\n") 
    return