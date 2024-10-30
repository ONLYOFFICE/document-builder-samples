import os
import argparse
import platform
import uuid

langs = ['cpp', 'cs']
os_name = platform.system().lower()

def mkdir(dir):
    if not os.path.exists(dir):
        os.mkdir(dir)

def log(level, message):
    print('configure.py: ' + level + ': ' + message)

def getDefaultBuilderDir():
    dir = ''
    if os_name == 'windows':
        dir = 'C:\\Program Files\\ONLYOFFICE\\DocumentBuilder'
    elif os_name == 'linux':
        dir = '/opt/onlyoffice/documentbuilder'
    return dir

def getAllTests():
    tests = {}
    for lang in langs:
        tests[lang] = []
        test_dirs = os.listdir(lang)
        for test_dir in test_dirs:
            if not os.path.isdir(lang + '/' + test_dir):
                continue
            tests[lang].append(lang + '/' + test_dir)
    return tests

def printAvailableTests():
    all_tests = getAllTests()
    print('all')
    for lang in langs:
        print('-----')
        print(lang)
        tests = all_tests[lang]
        for test in tests:
            print(test)

class PrintTestsList(argparse.Action):
    def __call__(self, parser, namespace, values, option_string=None):
        printAvailableTests()
        exit()

def getSelectedTests(tests):
    all_tests = getAllTests()
    # make set of all available tests
    available_tests = {'all'}
    for lang in langs:
        available_tests.add(lang)
        available_tests.update(all_tests[lang])
    # make dict with set of selected tests
    tests_selected = {lang: set() for lang in langs}
    # filter tests through only available ones
    for test in tests:
        if not test in available_tests:
            log('warning', 'wrong test "' + test + '". Call script with --list (or -l) to see all available tests')
            continue

        if test == 'all':
            for lang in langs:
                tests_selected[lang].update(all_tests[lang])
        elif not '/' in test:
            lang = test
            tests_selected[lang].update(all_tests[lang])
        else:
            lang = test.split('/')[0]
            tests_selected[lang].add(test)
    # delete empty tests
    for lang in langs:
        if not tests_selected[lang]:
            del tests_selected[lang]

    return tests_selected

def replacePlaceholders(template_file, output_file, replacements):
    content = ''
    # open and read template file
    with open(template_file, 'r') as file:
        content = file.read()
    # replace all placeholders with corresponding values
    for placeholder, replacement in replacements.items():
        content = content.replace(placeholder, replacement)
    # write result to output file
    with open(output_file, 'w') as file:
        file.write(content)

def genVSProjectsCPP(tests, builder_dir):
    if os_name != 'windows':
        log('warning', 'generating Visual Studio C++ projects is only available on Windows')
        return

    builder_dir = builder_dir.replace('/', '\\')
    for test in tests:
        test_dir = 'out/' + test
        mkdir(test_dir)
        test_name = test.split('/')[1]
        if os.path.exists(test_dir + '/' + test_name + '.vcxproj'):
            continue

        log('info', 'generating VS C++ project for sample "' + test + '"...')
        # .vcxproj
        project_guid = str(uuid.uuid4())
        replacements = {
            '[PROJECT_GUID]': project_guid,
            '[TEST_NAME]': test_name,
            '[BUILDER_DIR]': builder_dir,
            '[ROOT_DIR]': os.getcwd()
        }
        replacePlaceholders('configure/templates/cpp/template.vcxproj', test_dir + '/' + test_name + '.vcxproj', replacements)
        # .sln
        replacements = {
            '[SOLUTION_GUID]': str(uuid.uuid4()).upper(),
            '[TEST_NAME]': test_name,
            '[PROJECT_GUID]': project_guid.upper(),
            '[EXT_GLOBALS_GUID]': str(uuid.uuid4()).upper()
        }
        replacePlaceholders('configure/templates/cpp/template.sln', test_dir + '/' + test_name + '.sln', replacements)
        # .vcxproj.filters
        replacements = {
            '[GUID_SOURCE_FILES]': str(uuid.uuid4()).upper(),
            '[GUID_HEADER_FILES]': str(uuid.uuid4()).upper(),
            '[GUID_RESOURCE_FILES]': str(uuid.uuid4()).upper(),
            '[TEST_NAME]': test_name
        }
        replacePlaceholders('configure/templates/cpp/template.vcxproj.filters', test_dir + '/' + test_name + '.vcxproj.filters', replacements)
        # .vcxproj.user
        replacements = {
            '[BUILDER_DIR]': builder_dir
        }
        replacePlaceholders('configure/templates/cpp/template.vcxproj.user', test_dir + '/' + test_name + '.vcxproj.user', replacements)

def genVSProjectsCS(tests, builder_dir):
    if os_name != 'windows':
        log('warning', 'generating Visual Studio C# projects is only available on Windows')
        return

    builder_dir = builder_dir.replace('/', '\\')
    for test in tests:
        test_dir = 'out/' + test
        mkdir(test_dir)
        test_name = test.split('/')[1]
        if os.path.exists(test_dir + '/' + test_name + '.csproj'):
            continue

        log('info', 'generating VS C# project for sample "' + test + '"...')
        # .csproj
        project_guid = str(uuid.uuid4())
        replacements = {
            '[TEST_NAME]': test_name,
            '[BUILDER_DIR]': builder_dir,
        }
        replacePlaceholders('configure/templates/cs/template.csproj', test_dir + '/' + test_name + '.csproj', replacements)
        # .sln
        replacements = {
            '[SOLUTION_GUID]': str(uuid.uuid4()).upper(),
            '[TEST_NAME]': test_name,
            '[PROJECT_GUID]': project_guid.upper(),
            '[EXT_GLOBALS_GUID]': str(uuid.uuid4()).upper()
        }
        replacePlaceholders('configure/templates/cs/template.sln', test_dir + '/' + test_name + '.sln', replacements)

def genQtProjectsCPP(tests, builder_dir):
    root_dir = os.getcwd()
    if os_name == 'windows':
        builder_dir = builder_dir.replace('\\', '/')
        root_dir = root_dir.replace('\\', '/')

    for test in tests:
        test_dir = 'out/' + test
        mkdir(test_dir)
        test_name = test.split('/')[1]
        if os.path.exists(test_dir + '/' + test_name + '.pro'):
            continue

        log('info', 'generating Qt C++ project for sample "' + test + '"...')
        # .pro
        replacements = {
            '[TEST_NAME]': test_name,
            '[BUILDER_DIR]': builder_dir,
            '[ROOT_DIR]': root_dir
        }
        replacePlaceholders('configure/templates/cpp/template.pro', test_dir + '/' + test_name + '.pro', replacements)

def genMakefileCPP(tests, builder_dir):
    if os_name == 'windows':
        log('warning', 'generating Makefile is not available on Windows')
        return

    # initialize variables
    compiler = ''
    lflags = ''
    env_lib_path = ''
    if os_name == 'linux':
        compiler = 'g++'
        lflags = '-Wl,--unresolved-symbols=ignore-in-shared-libs'
        env_lib_path = 'LD_LIBRARY_PATH'
    elif os_name == 'darwin':
        compiler = 'clang++'
        env_lib_path = 'DYLD_LIBRARY_PATH'
    root_dir = os.getcwd()
    for test in tests:
        test_dir = 'out/' + test
        mkdir(test_dir)
        test_name = test.split('/')[1]
        if os.path.exists(test_dir + '/Makefile'):
            continue

        log('info', 'generating Makefile for C++ sample "' + test + '"...')
        # Makefile
        replacements = {
            '[TEST_NAME]': test_name,
            '[BUILDER_DIR]': builder_dir,
            '[ROOT_DIR]': root_dir,
            '[COMPILER]': compiler,
            '[LFLAGS]': lflags,
            '[ENV_LIB_PATH]': env_lib_path
        }
        replacePlaceholders('configure/templates/cpp/Makefile', test_dir + '/Makefile', replacements)

def genCPP(projects, tests, builder_dir):
    mkdir('out/cpp')
    # generate header with builder path
    if not os.path.exists('out/cpp/builder_path.h'):
        replacements = {
            '[BUILDER_DIR]': builder_dir.replace('\\', '/')
        }
        replacePlaceholders('configure/templates/cpp/builder_path.h', 'out/cpp/builder_path.h', replacements)
    # VS
    if projects['vs']:
        genVSProjectsCPP(tests, builder_dir)
    # Qt
    if projects['qt']:
        genQtProjectsCPP(tests, builder_dir)
    # Makefile
    if projects['make']:
        genMakefileCPP(tests, builder_dir)

def genCS(projects, tests, builder_dir):
    mkdir('out/cs')
    # generate file with builder path
    if not os.path.exists('out/cs/Constants.cs'):
        replacements = {
            '[BUILDER_DIR]': builder_dir.replace('\\', '/')
        }
        replacePlaceholders('configure/templates/cs/Constants.cs', 'out/cs/Constants.cs', replacements)
    # VS
    if projects['vs']:
        genVSProjectsCS(tests, builder_dir)
    else:
        log('warning', 'generating C# projects only available ' + ('on Windows ' if os_name != 'windows' else '') + 'with --vs')


if __name__ == '__main__':
    # go to root dir
    file_dir = os.path.dirname(os.path.realpath(__file__))
    os.chdir(file_dir + '/..')
    # initialize argument parser
    parser = argparse.ArgumentParser(description='Generate project files for Document Builder samples')
    parser.add_argument('--vs', action='store_true', help='create Visual Studio (.vcxproj and .csproj) project files')
    parser.add_argument('--qt', action='store_true', help='create Qt (.pro) project files')
    parser.add_argument('--make', action='store_true', help='create Makefile')
    parser.add_argument('-t', '--test', dest='tests', action='append', help='specifies tests to generate project files', required=True)
    parser.add_argument('-l', '--list', action=PrintTestsList, nargs=0, help='show list of available tests and exit')

    default_builder_dir = getDefaultBuilderDir()
    if default_builder_dir:
        parser.add_argument('--dir', action='store', help='specifies Document Builder directory (default: ' + default_builder_dir + ')', default=default_builder_dir)
    else:
        parser.add_argument('--dir', action='store', help='specifies Document Builder directory', required=True)

    args = parser.parse_args()

    # validate arguments
    if not os.path.exists(args.dir):
        log('error', 'Document Builder directory doesn\'t exist: ' + args.dir)
        exit(1)

    if not (args.vs or args.qt or args.make):
        if os_name == 'windows':
            args.vs = True
        args.qt = True
        if os_name != 'windows':
            args.make = True

    projects = {'vs': args.vs, 'qt': args.qt, 'make': args.make}
    # filter tests
    tests_selected = getSelectedTests(args.tests)
    # generate project files
    mkdir('out')
    handlers = {'cpp': genCPP, 'cs': genCS}
    for lang, tests in tests_selected.items():
        handlers[lang](projects, tests, args.dir)
