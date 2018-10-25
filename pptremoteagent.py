from win32gui import GetWindowText, GetForegroundWindow
import time, requests, sys, getopt, keyboard

REQUESTS_CONNECT_TIMEOUT    = 5
REQUESTS_READ_TIMEOUT       = 5

COMMANDS = ['next', 'back', 'stop']

def getcommand (p_server):
    if True:
        r = requests.get('http://' + p_server + '/command', timeout=(REQUESTS_CONNECT_TIMEOUT, REQUESTS_READ_TIMEOUT))
    else:
        print('WARNING: unable to access command server')
        return(None)
        
    if r.status_code != requests.codes.ok:
        return (None)
        
    rjson = r.json()
    
    if not rjson is None:
        rcmd = rjson['command']
        if rcmd in COMMANDS:
            return(rcmd)
    
    return(None)


def main(argv):
    #initialize variables for command line arguments
    arg_server = '127.0.0.1:5000'

    #get command line arguments
    try:
        opts, args = getopt.getopt(argv, 's:')
    except getopt.GetoptError:
        printhelp()
        sys.exit(2)
        
    for opt, arg in opts:
        if opt == '-s':
            arg_server = arg

    print ('INFO: Polling server %s' % arg_server)        
            
    while (1):
        time.sleep(1)
        wintext = GetWindowText(GetForegroundWindow())
                
        if wintext.startswith('PowerPoint Slide Show - ['):
            cmd = getcommand (arg_server)
            if not cmd is None:
                if   cmd == 'next':
                    keyboard.send('space')
                elif cmd == 'back':
                    keyboard.send('backspace')
                elif cmd == 'stop':
                    keyboard.send('escape')
            
if __name__ == '__main__':
    main(sys.argv[1:])            
