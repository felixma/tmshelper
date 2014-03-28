#!/usr/bin/env python

"""tmshelper is a tool to extract information from MS excel file into TMS staff.

History:

3/28/2014 - version 0.1
"""
__authors__ = ['"Felix Ma" <felix.ma@alcatel-lucent.com>']

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


#def usage():
#    print 'tmshelper usage:'
#    print '-h, --help: print help message.'
#    print '-i: Test plan excel file.'
#    print '-a: Generate all (schedule, plan, script)'
#    print '-s: Generate schedule only'
#    print '-p: Generate plan only'
#    print '-t: Generate script only'

def getinfo(testplan):
    """Extract general information (e.g., setname, grp) from MS excel file, initialize global parameters.

    Args:
        testplan: Prepared MS excel file name

    Raises:
        IOError: When testplan doesn't exist.
    """
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
    """Generate TMS schedule file.

    Args:
        None.

    Output:
        sched_grp (grp is defined in MS excel file.)
    """
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
    """Generate TMS test plan file.

    Args:
        testplan: Prepared MS excel file name

    Raises:
        IOError: When testplan doesn't exist.

    Output:
        plan_grp (grp is defined in MS excel file.)
    """
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
    """Generate TMS test script file.

    Args:
        testplan: Prepared MS excel file name

    Raises:
        IOError: When testplan doesn't exist.
        AttributeError: When regular expression search fails.

    Output:
        script_grp/script_files (grp is defined in MS excel file.)
    """
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
    """Parse command line arguments.

    Args:
        Command line arguments.

    Raise:
        TypeError: When no additional command line argument.
    """
    parser = optparse.OptionParser()
    parser.add_option('-i', action = 'store', dest = 'testplan', help = 'test plan excel file')
    parser.add_option('-a', action = 'store_true', dest = 'all', help = 'generate all (schedule, plan, script)')
    parser.add_option('-s', action = 'store_true', dest = 'schedOnly', help = 'generate schedule only')
    parser.add_option('-p', action = 'store_true', dest = 'planOnly', help = 'generate plan only')
    parser.add_option('-t', action = 'store_true', dest = 'scriptOnly', help = 'generate script only')
    opt, arg = parser.parse_args(argv)
    
    testplan = opt.testplan
    getinfo(testplan)

    #if opt.all or (opt.all is None and opt.schedOnly is None and opt.planOnly is None and opt.scriptOnly is None):
    if opt.all or not (opt.all or  opt.schedOnly or opt.planOnly or opt.scriptOnly):
        #print 'creating all'
        schedule()
        plan(testplan)
        script(testplan)
        sys.exit(0)
    if opt.schedOnly:
        #print 'creating schedule only'
        schedule()
    if opt.planOnly:
        #print 'creating plan only'
        plan(testplan)
    if opt.scriptOnly:
        #print 'creating script only'
        script(testplan)

if __name__ == '__main__':
    main(sys.argv[1:])

