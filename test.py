import os,sys,time

for i in range(5):
    time.sleep(1)
    sys.stdout.write('{} \r'.format(i))
    sys.stdout.flush()

exit(0)