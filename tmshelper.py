#!/usr/bin/env python

import os,shutil,re,sys,optparse
import xlrd

setname = ''
grp = ''
begindate = ''
enddate = ''
totalCaseNum = ''
originator = ''
rload = ''
pload = ''
RCR = ''
featureName = ''

def usage():
    print 'tmshelper usage:'
    print '-h, --help: print help message.'
    print '-i: Test plan excel file.'
    print '-a: Generate all (schedule, plan, script)'
    print '-s: Generate schedule only'
    print '-p: Generate plan only'
    print '-t: Generate script only'

def getinfo(testplan):
    with xlrd.open_workbook(testplan) as all:
        info = all.sheet_by_name(u'Info')
        global setname, grp,begindate, enddate, totalCaseNum, originator, rload, pload, RCR, featureName
        setname = info.cell(0,1).value
        grp = info.cell(1,1).value
        begindate = info.cell(2,1).value
        enddate = info.cell(3,1).value
        totalCaseNum = info.cell(4,1).value
        originator = info.cell(5,1).value
        rload = info.cell(6,1).value
        pload = info.cell(7,1).value
        RCR = info.cell(8,1).value
        featureName = info.cell(9,1).value

def schedule():
    with open('sched_%s' % grp, 'w') as sched:
        sched.write('setname = %s\n' % setname)
        sched.write('apple = IMS\n')
        sched.write('grp = %s\n' % grp)
        sched.write('phase = a\n')
        sched.write('testorg =\n')
        sched.write('tteam = cy\n')
        sched.write('begindate = %s\n' % begindate)
        sched.write('enddate = %s\n' % enddate)
        sched.write('planwritten = %s\n' % totalCaseNum)
        sched.write('planrun = %s\n' % totalCaseNum)
        sched.write('pcomment =\n')

def plan(testplan):
    with xlrd.open_workbook(testplan) as all:
        su = all.sheet_by_name(u'SU')
        regression = all.sheet_by_name(u'Regression')
        fix = all.sheet_by_name(u'AR_Defect')
        feat = all.sheet_by_name(u'Feature')
        tid = []
        for table in [su, regression, fix, feat]:
            t = table.col_values(0)
            if len(t) > 1:  # the first column is 'TID'
                t.pop(0)
                tid.extend(t)
        with open('plan_%s' % grp, 'w') as plan:
            for i in range(len(tid)):
                plan.write('setname = %s\n' % setname)
                plan.write('appl = IMS\n')
                plan.write('grp = %s\n' % grp)
                plan.write('phase = a\n')
                plan.write('tidnum = %s\n' % tid[i])
                plan.write('pload = %s\n' % pload)
                plan.write('pplace = SYSLAB\n')
                plan.write('testtype = NEW\n')
                plan.write('pusr1 = \n')
                plan.write('pusr2 = func\n')
                plan.write('plancase = \n')
                plan.write('pconfig = p1\n')
                plan.write('pparams = \n')
                plan.write('pcomment = \n')
                plan.write('\n')

def script(testplan):
    with xlrd.open_workbook(testplan) as all:
        scriptDirName = 'script_%s' % grp
        if os.path.exists(scriptDirName):
            shutil.rmtree(scriptDirName)
        os.mkdir(scriptDirName)
        os.chdir(scriptDirName)
        su = all.sheet_by_name(u'SU')
        regression = all.sheet_by_name(u'Regression')
        fix = all.sheet_by_name(u'AR_Defect')
        feat = all.sheet_by_name(u'Feature')
        #pattern = r'[\w\s\-\.\#\*"\'\?\\:/]*'
        for table in [su, regression, fix, feat]:
            t = table.col_values(0)
            s = table.col_values(1)
            if len(t) > 1:  # the first column is 'TID'
                t.pop(0)
                s.pop(0)
                for i in range(len(t)):
                    with open('%s.s' % t[i], 'w') as scr:
                        scr.write('Test I.D.: %s\n' % t[i])
                        
                        reg_title = r'Title:([\w\s\-]*)Initial'
                        title = re.search(reg_title, s[i]).group(1).strip()
                        scr.write('Title: %s\n' % title)

                        scr.write('Owner: C.YU\n')
                        scr.write('Originator: %s\n' % originator)
                        scr.write('Script Status: I\n')
                        if table is not feat:
                            scr.write('Requirement(s): Implicit\n')
                            scr.write('Feature(s): Implicit\n')
                        else:
                            scr.write('Requirement(s): RCR%s\n' % RCR)
                            scr.write('Feature(s): %s\n' % featureName) 
                        scr.write('Reference(s):\n')
                        scr.write('Functional Area(s): 90\n')
                        scr.write('Test Level: G\n')
                        scr.write('Original Target Application(s): IMS\n')
                        scr.write('Parent Test(s):\n')
                        scr.write('Execution Mode: MAN\n')
                        scr.write('Estimated Execution Time: 30\n')
                        scr.write('\n')
                        scr.write('Description:\n')
                        scr.write('%s\n' % title)
                        scr.write('\n')
                        scr.write('Issue Notes:\nnone\n\n')
                        scr.write('Resources/Configuration:\nnone\n\n')
                        
                        scr.write('Initial Conditions/System Setup:\n')
                        reg_config = r'Initial Configuration:([\w\s\-\.\#\*"\'\?\\:/]*)Test Procedure'
                        configstr = re.search(reg_config, s[i])
                        if configstr is not None:
                            config = configstr.group(1).strip()
                            scr.write('%s\n' % config)

                        reg_procedure = r'Test Procedure:([\w\s\-\.\#\*"\'\?\\:/]*)Verify'
                        procedure = re.search(reg_procedure, s[i]).group(1).strip()
                        scr.write('Test Procedure:\n%s\n\n' % procedure)

                        reg_verify = r'Verify:([\w\s\-\.\#\*"\'\?\\:/]*)'
                        verifystr = re.search(reg_verify, s[i])
                        if verifystr:
                            scr.write('Verify:\n%s\n' % verifystr.group(1).strip())
                        else:
                            scr.write('Verify:\n')
                        


def main(argv):
    parser = optparse.OptionParser()
    parser.add_option('-i', action = 'store', dest = 'testplan')
    opt, arg = parser.parse_args(argv)

    getinfo(opt.testplan)
    schedule()
    plan(opt.testplan)
    script(opt.testplan)

if __name__ == '__main__':
    main(sys.argv[1:])

