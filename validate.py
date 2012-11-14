#!/usr/bin/env python

#validate output files as much as possible
import os
import sys
import time

#list all the files in mods_files
for f in os.listdir('mods_files'):
    #run xmllint on each one
    filename = './mods_files/' + f
    print('Validating %s...' % f)
    #xmllint --noout --schema "http://www.loc.gov/standards/mods/v3/mods-3-4.xsd" 00001817.mods
    #args = ['/usr/bin/xmllint',  '--noout', '--schema', 
    #    './mods-3-4.xsd.nice', './mods_files/' + f]
    #p = Popen(args)
    cmd = '/usr/bin/xmllint --noout --schema "./mods-3-4.xsd" %s' % filename
    result = os.system(cmd)
    if result != 0:
        print 'Error'
    time.sleep(3)

sys.exit()

