# Assuming you have not changed the general structure of the template no modification is needed in this file.
from . import commands
from .lib import fusionAddInUtils as futil
import threading
import time
import platform
import os
from datetime import datetime
from configparser import ConfigParser
import subprocess
from subprocess import STDOUT, PIPE
from zipfile import ZipFile
import json
import sys
import re
import xml.etree.ElementTree as ET

try:
    from urllib2 import Request, urlopen, HTTPError
except ImportError:
    from urllib.request import Request, urlopen
    from urllib.error import HTTPError

VERSION = "1.0.0"

handlers = []
lastHeartbeat = 0
is_win = platform.system() == 'Windows'
pallete = None

HEARTBEAT_INTERVAL = 180
API_KEY = ""
API_URL = "https://api.wakatime.com/api/v1/users/current/heartbeats"
HOME_FOLDER = os.path.realpath(os.environ.get('WAKATIME_HOME') or os.path.expanduser('~'))
CONFIG_FILE = os.path.join(HOME_FOLDER, '.wakatime.cfg')
RESOURCES_FOLDER = os.path.join(HOME_FOLDER, '.wakatime')
WAKATIME_CLI_LOCATION = None
GITHUB_RELEASES_STABLE_URL = 'https://api.github.com/repos/wakatime/wakatime-cli/releases/latest'
GITHUB_DOWNLOAD_PREFIX = 'https://github.com/wakatime/wakatime-cli/releases/download'
INTERNAL_CONFIG_FILE = os.path.join(HOME_FOLDER, '.wakatime-internal.cfg')
LATEST_CLI_VERSION = None

class Popen(subprocess.Popen):
    """Patched Popen to prevent opening cmd window on Windows platform."""

    def __init__(self, *args, **kwargs):
        if is_win:
            startupinfo = kwargs.get('startupinfo')
            try:
                startupinfo = startupinfo or subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            except AttributeError:
                pass
            kwargs['startupinfo'] = startupinfo
        super(Popen, self).__init__(*args, **kwargs)

class Wakatime:
    def __init__(self):
        self.currentFile = None
        self.currentProject = None
        self.language = "Autodesk Fusion 360"
        self.category = "designing"
        self.configs = None

    def loadConfig(self):
        global API_KEY, API_URL, CONFIG_FILE

        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as configFile:
                configs = None
                try:
                    configs = parseConfigFile(CONFIG_FILE)
                    self.configs = configs
                    if configs:
                        if configs.has_option('settings', 'api_url'):
                            url = configs.get('settings', 'api_url')
                            if url:
                                API_URL = url.strip()
                        if configs.has_option('settings', 'api_key'):
                            key = configs.get('settings', 'api_key')
                            if key:
                                API_KEY = key.strip()
                            else:
                                result, cancelled = futil.ui.inputBox("Enter your Wakatime API Key: ", "Wakatime Setup", "")
                                if result and not cancelled:
                                    API_KEY = result.strip()

                                    if not os.path.exists(os.path.dirname(CONFIG_FILE)):
                                        os.makedirs(os.path.dirname(CONFIG_FILE))
                                    with open(CONFIG_FILE, 'w') as configFile:
                                        configs.set('settings', 'api_key', API_KEY)
                                        configs.write(configFile)
                except Exception as e:
                    futil.handle_error(f'Error loading API key: {e}')

    def getCurrentFileInfo(self):
        try:
            doc = futil.app.activeDocument
            
            if doc:
                localFile = futil.app.executeTextCommand("Document.Path").replace('"', '')
                try:
                    metaFile = localFile + '._xx' 
                    tree = ET.parse(metaFile)
                    root = tree.getroot()
                    for prN in root.iter('ProjectName'):
                        projectName = prN.text
                except:
                    projectName = "Unknown"
                return localFile, projectName

        except Exception as e:
            futil.handle_error(f'Error getting current file info: {e}')
            pass
        
        return "Untitled", "Unknown"

    def sendHeartbeat(self, isWrite=False):
        global API_KEY, lastHeartbeat, stopTracking

        currentTime = time.time()
        if currentTime - lastHeartbeat < HEARTBEAT_INTERVAL:
            return

        lastHeartbeat = time.time()           


        entity, project = self.getCurrentFileInfo()

        cmd = [
            getCliLocation(),
            '--entity', entity,
            '--project', project,
            '--language', self.language,
            '--category', self.category,
            '--time', str(datetime.now().timestamp()),
            '--plugin', 'Fusion360/{0} Fusion360Wakatime/{1}'.format(futil.app.version, VERSION),
        ]

        if API_KEY:
            cmd.extend(['--key', API_KEY])
        if API_URL:
            cmd.extend(['--api-url', API_URL])

        if isWrite:
            cmd.append('--write')

        try:
            process = Popen(cmd, stdout=PIPE, stderr=STDOUT)
            output, err = process.communicate()
            retcode = process.poll()

            if (not retcode or retcode == 102 or retcode == 112) and not output:
                futil.app.log('Heartbeat sent successfully.')
            if retcode:
                futil.app.log('wakatime-core exited with status: {0}'.format(retcode))
            if output:
                futil.app.log('wakatime-core output: {0}'.format(output))
        except Exception as e:
            futil.handle_error(f'Error sending heartbeat: {e}')
            return

def request(url, last_modified=None):
    req = Request(url)
    req.add_header('User-Agent', 'github.com/itzshubhamdev/fusion360-wakatime')

    if last_modified:
        req.add_header('If-Modified-Since', last_modified)

    try:
        resp = urlopen(req)
        headers = dict(resp.getheaders()) if is_py2 else resp.headers
        return headers, resp.read(), resp.getcode()
    except HTTPError as err:
        if err.code == 304:
            return None, None, 304
        if is_py2:
            with SSLCertVerificationDisabled():
                try:
                    resp = urlopen(req)
                    headers = dict(resp.getheaders()) if is_py2 else resp.headers
                    return headers, resp.read(), resp.getcode()
                except HTTPError as err2:
                    if err2.code == 304:
                        return None, None, 304
                    raise
                except IOError:
                    raise
        raise
    except IOError:
        if is_py2:
            with SSLCertVerificationDisabled():
                try:
                    resp = urlopen(url)
                    headers = dict(resp.getheaders()) if is_py2 else resp.headers
                    return headers, resp.read(), resp.getcode()
                except HTTPError as err:
                    if err.code == 304:
                        return None, None, 304
                    raise
                except IOError:
                    raise
        raise

def download(url, filePath):
    req = Request(url)
    req.add_header('User-Agent', 'github.com/itzshubhamdev/fusion360-wakatime')

    with open(filePath, 'wb') as fh:
        try:
            resp = urlopen(req)
            fh.write(resp.read())
        except HTTPError as err:
            if err.code == 304:
                return None, None, 304
            if is_py2:
                with SSLCertVerificationDisabled():
                    try:
                        resp = urlopen(req)
                        fh.write(resp.read())
                        return
                    except HTTPError as err2:
                        raise
                    except IOError:
                        raise
            raise
        except IOError:
            if is_py2:
                with SSLCertVerificationDisabled():
                    try:
                        resp = urlopen(url)
                        fh.write(resp.read())
                        return
                    except HTTPError as err:
                        raise
                    except IOError:
                        raise
            raise

def updateCli():
    if isCliLatest():
        return False

    if os.path.isdir(os.path.join(RESOURCES_FOLDER, 'wakatime-cli')):
        shutil.rmtree(os.path.join(RESOURCES_FOLDER, 'wakatime-cli'))

    if not os.path.exists(RESOURCES_FOLDER):
        os.makedirs(RESOURCES_FOLDER)

    try:
        url = cliDownloadUrl()
        zip_file = os.path.join(RESOURCES_FOLDER, 'wakatime-cli.zip')
        download(url, zip_file)

        if isCliInstalled():
            try: 
                os.remove(getCliLocation())
            except:
                futil.handle_error('Failed to remove old Wakatime CLI binary.')

        with ZipFile(zip_file) as zf:
            zf.extractall(RESOURCES_FOLDER)

        if not is_win:
            os.chmod(getCliLocation(), 509) # 755

        try:
            os.remove(os.path.join(RESOURCES_FOLDER, 'wakatime-cli.zip'))
        except:
            futil.handle_error('Failed to remove downloaded Wakatime CLI zip file.')

    except Exception as e:
        futil.handle_error(f'Error updating Wakatime CLI: {e}')

def getCliLocation():
    global WAKATIME_CLI_LOCATION

    if not WAKATIME_CLI_LOCATION:
        binary = 'wakatime-cli-{osname}-{arch}{ext}'.format(
            osname=platform.system().lower(),
            arch=architecture(),
            ext='.exe' if is_win else '',
        )
        WAKATIME_CLI_LOCATION = os.path.join(RESOURCES_FOLDER, binary)

    return WAKATIME_CLI_LOCATION

def architecture():
    arch = platform.machine() or platform.processor()
    if arch == 'armv7l':
        return 'arm'
    if arch == 'aarch64':
        return 'arm64'
    if 'arm' in arch:
        return 'arm64' if sys.maxsize > 2**32 else 'arm'
    return 'amd64' if sys.maxsize > 2**32 else '386'

def isCliInstalled():
    return os.path.exists(getCliLocation())

def isCliLatest():
    if not isCliInstalled():
        return False

    args = [getCliLocation(), '--version']

    try:
        stdout, stderr = Popen(args, stdout=PIPE, stderr=PIPE).communicate()
    except Exception as e:
        return False

    stdout = (stdout or b'') + (stderr or b'')
    localVer = extractVersion(stdout.decode('utf-8'))
    if not localVer:
        return False

    remoteVer = getLatestCliVersion()

    if not remoteVer:
        return True

    if remoteVer == localVer:
        return True

    return False

def extractVersion(text):
    pattern = re.compile(r"([0-9]+\.[0-9]+\.[0-9]+)")
    match = pattern.search(text)
    if match:
        return 'v{ver}'.format(ver=match.group(1))
    return None

def getLatestCliVersion():
    global LATEST_CLI_VERSION

    if LATEST_CLI_VERSION:
        return LATEST_CLI_VERSION

    configs, last_modified, last_version = None, None, None
    try:
        configs = parseConfigFile(INTERNAL_CONFIG_FILE)
        if configs:
            last_modified, last_version = lastModifiedAndVersion(configs)
    except Exception as e:
        futil.handle_error(f'Error reading internal config file: {e}')

    try:
        headers, contents, code = request(GITHUB_RELEASES_STABLE_URL, last_modified=last_modified)

        if code == 304:
            LATEST_CLI_VERSION = last_version
            return last_version

        data = json.loads(contents.decode('utf-8'))

        ver = data['tag_name']

        if configs:
            last_modified = headers.get('Last-Modified')
            if not configs.has_section('internal'):
                configs.add_section('internal')
            configs.set('internal', 'cli_version', ver)
            configs.set('internal', 'cli_version_last_modified', last_modified)
            with open(INTERNAL_CONFIG_FILE, 'w', encoding='utf-8') as fh:
                configs.write(fh)

        LATEST_CLI_VERSION = ver
        return ver
    except:
        return None

def lastModifiedAndVersion(configs):
    last_modified, last_version = None, None
    if configs.has_option('internal', 'cli_version'):
        last_version = configs.get('internal', 'cli_version')
    if last_version and configs.has_option('internal', 'cli_version_last_modified'):
        last_modified = configs.get('internal', 'cli_version_last_modified')
    if last_modified and last_version and extractVersion(last_version):
        return last_modified, last_version
    return None, None

def cliDownloadUrl():
    osname = platform.system().lower()
    arch = architecture()

    validCombinations = [
      'darwin-amd64',
      'darwin-arm64',
      'freebsd-386',
      'freebsd-amd64',
      'freebsd-arm',
      'linux-386',
      'linux-amd64',
      'linux-arm',
      'linux-arm64',
      'netbsd-386',
      'netbsd-amd64',
      'netbsd-arm',
      'openbsd-386',
      'openbsd-amd64',
      'openbsd-arm',
      'openbsd-arm64',
      'windows-386',
      'windows-amd64',
      'windows-arm64',
    ]
    check = '{osname}-{arch}'.format(osname=osname, arch=arch)

    version = getLatestCliVersion()

    return '{prefix}/{version}/wakatime-cli-{osname}-{arch}.zip'.format(
        prefix=GITHUB_DOWNLOAD_PREFIX,
        version=version,
        osname=osname,
        arch=arch,
    )

def parseConfigFile(configFile):
    configs = ConfigParser()
    try:
        with open(configFile, 'r', encoding='utf-8') as fh:
            try:
                configs.read_file(fh)
                return configs
            except ConfigParserError:
                return None
    except IOError:
        return configs

class DocumentSavedHandler(futil.adsk.core.DocumentEventHandler):
    def __init__(self):
        super().__init__()
    
    def notify(self, args):
        try:
            tracker.sendHeartbeat(isWrite=True)
        except:
            futil.app.log('Error sending heartbeat on document save', futil.adsk.core.LogLevels.ErrorLogLevel)
            pass          

class ActionSelectionChangeHandler(futil.adsk.core.ActiveSelectionEventHandler):
    def __init__(self):
        super().__init__()
    
    def notify(self, args):
        try:
            tracker.sendHeartbeat()
        except:
            futil.app.log('Error sending heartbeat on selection change', futil.adsk.core.LogLevels.ErrorLogLevel)
            pass

def run(context):
    global pallete
    try:
        app = futil.adsk.core.Application.get()
        updateCli()

        onDocumentSaved = DocumentSavedHandler()
        app.documentSaved.add(onDocumentSaved)
        handlers.append(onDocumentSaved)

        ActiveSelectionChanged = ActionSelectionChangeHandler()
        app.userInterface.activeSelectionChanged.add(ActiveSelectionChanged)
        handlers.append(ActiveSelectionChanged)

        tracker.loadConfig()
        tracker.sendHeartbeat()
    except:
        futil.handle_error('run')

tracker = Wakatime()

def stop(context):
    try:
        futil.clear_handlers()
    except:
        futil.handle_error('stop')

def is_symlink(path):
    try:
        return os.is_symlink(path)
    except:
        return False


def createSymlink():
    link = os.path.join(RESOURCES_FOLDER, 'wakatime-cli')
    if is_win:
        link = link + '.exe'
    elif os.path.exists(link) and is_symlink(link):
        return  # don't re-create symlink on Unix-like platforms

    try:
        os.symlink(getCliLocation(), link)
    except:
        try:
            shutil.copy2(getCliLocation(), link)
            if not is_win:
                os.chmod(link, 509)  # 755
        except:
            futil.handle_error('Failed to create symlink for Wakatime CLI binary. Please create a symlink manually or copy the binary to the resources folder.')
            return


class SSLCertVerificationDisabled(object):

    def __enter__(self):
        self.original_context = ssl._create_default_https_context
        ssl._create_default_https_context = ssl._create_unverified_context

    def __exit__(self, *args, **kwargs):
        ssl._create_default_https_context = self.original_context