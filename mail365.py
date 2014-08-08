#!/usr/bin/env python

#   File:      mail365.py                                                                    
#   Purpose:   mail365 - an electronic mail transport agent for Office 365                   
#   Author:    Paul Greenberg (http://www.greenberg.pro)                                                                
#   Version:   1.0
#   Copyright: (c) 2014 Paul Greenberg <paul@greenberg.pro>
#
#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.
#
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#
#   You should have received a copy of the GNU General Public License
#   along with this program.  If not, see <http://www.gnu.org/licenses/>.


import os;
import pwd;
import sys;
if sys.version_info[0] < 3:
    sys.stderr.write(os.path.basename(__file__) + ' requires Python 3 or higher.\n');
    sys.stderr.write("python3 " + __file__ + '\n');
    exit(1);

import argparse;
import pprint;
import time;
import datetime;
import re;
from email.header import decode_header;
from pyewsclient import EWSSession, EWSEmail, EWSAttachment;

log_lvl = 0;
log = [];
err = [];
msg = '';
lc = 0;
usr = pwd.getpwuid(os.getuid())[0];
ews = {};
hdr = {};
bdy = {};


def _exit(lvl=0):
    if log_lvl > 0:
        if log:
            print('\n'.join(log));
    else:
        if err:
            print('\n'.join(err));
    if lvl == 1:
        exit(1);
    else:
        exit(0);


def _read_mail():
    global hdr, bdy;
    lnb = 0;
    ln = 0;
    loc = 'header';
    hn = '';
    hv = '';
    bndr = None;
    bi = 0;
    bib = 0;
    for line in msg.split('\n'):
        ln += 1;

        if line == '':
            lnb += 1;

        if lnb == 2:
            loc = 'body';
            if 'Content-Type' not in hdr:
                _log('Content-Type header does not exist', 'CRIT', 'ERR');
                _exit(1);

            m = re.match('multipart/mixed; boundary="(.*)"$', hdr['Content-Type']);
            if m:
                bndr = m.group(1);
            else:
                _log('Content-Type is not multipart/mixed or is missing boundary', 'CRIT', 'ERR');
                _exit(1);
            continue;

        if loc == 'header':
            if re.match('\w', line):
                m = re.match("([A-Za-z0-9-]+):\s(.*)", line);
                if m:
                    hn = m.group(1);
                    hdr[hn] = m.group(2);
            else:
                hdr[hn] += line.strip();

        if loc == 'body':
            if bndr is None:
                _log('Boundary is not defined', 'CRIT', 'ERR');
                _exit(1);

            if re.match('--' + bndr + '--$', line):
               break;

            if re.match('--' + bndr + '$', line):
                bi += 1;
                bib = 0;
                try:
                    f = bdy[bi];
                except:
                    bdy[bi] = {};
                continue;

            if bib == 1:
                 if 'Body' not in bdy[bi]:
                     bdy[bi]['Body'] = [];
                 bdy[bi]['Body'].append(line);
                
            if bib == 0:
                m = re.match("([A-Za-z0-9-]+):\s(.*)", line);
                if m:
                    bdy[bi][m.group(1)] = m.group(2);
                if line == '':
                    bib = 1;
             
    return;


def _read_pipe():
    global lc, msg;
    lc = 0;
    b = [];
    if sys.stdin.isatty():
        _log('improper usage. must redirect e-mail message to this script using IO pipes', 'CRIT', 'ERR');
        _exit(1);
    else:
        while True:
            c = sys.stdin.read(1);
            if c == '':
                break;
            else:
                b.append(c);
                lc += 1;
    msg = ''.join(b);
    return;


def _load_conf():
    global ews, usr;
    conf_file = '.mail365.conf';
    if not os.path.isfile(os.path.expanduser('~') + '/' + conf_file):
        _log('configuration file ~/' + conf_file + ' does not exist', 'CRIT', 'ERR');
        _exit(1);
    else:
        _log('username: ' + usr, 'INFO', 'LOG');
        _log('configuration file ~/' + conf_file + ' exists', 'INFO', 'LOG');
        c = None;
        with open(os.path.expanduser('~') + '/' + conf_file, 'r') as f:
            c = f.read();
        if c is not None:
            for i in c.split('\n'):
                m = re.match('([A-Za-z0-9_]+)=(.*)', i);
                if m:
                    ews[m.group(1)] = m.group(2);
    return;


def _log(msg='TEST', lvl='INFO', fcl='LOG'):
    global log_lvl, log, err;
    ''' Log Processing '''
    cls = str(sys.__name__);
    func = str(sys._getframe(1).f_code.co_name);
    ts = str(datetime.datetime.now());
    for xmsg in msg.split('\n'):
        if fcl == 'ERR':
            err.append("{0} | {1} | {2} | {3}".format(ts, cls + '.' + func + '()', fcl + '.' + lvl, xmsg));
        log.append("{0:26s} | {1:45s} | {2:7s} | {3}".format(ts, cls + '.' + func + '()', fcl + '.' + lvl, xmsg));
    return;


def main():
    global log_lvl;
    func = 'main()';
    parser = argparse.ArgumentParser(description='mail365 - an electronic mail transport agent for Office 365');
    parser.add_argument('-l', '--log', dest='ilog', metavar='LOGLEVEL', type=int, help='log level (default: 0)');
    parser.add_argument('-t', '--trust', dest='itrust', metavar='TRUST', required=False, help='Read message for recipients. To:, Cc:, and Bcc: lines will be scanned for recipient addresses. The Bcc: line will be deleted before transmission.');
    args = parser.parse_args();

    if args.ilog:
        log_lvl = args.ilog;

    _read_pipe();
    _read_mail();
    _load_conf();

    ebdy = '';
    esbj = '';
    eto = [];
    ecc = [];
    ebcc = [];

    if 'Date' in hdr:
        ebdy += 'Date: ' + hdr['Date'] + '\n';

    if 'From' in hdr:
        ebdy += 'From: ' + hdr['From'] + '\n';

    if 'To' in hdr:
        ebdy += 'To: ' + hdr['To'] + '\n';
        for i in hdr['To'].split(';'):
            eto.append(i);

    if 'Cc' in hdr:
        ebdy += 'Cc: ' + hdr['Cc'] + '\n';
        for i in hdr['Cc'].split(';'):
            ecc.append(i);

    if 'Bcc' in hdr:
        ebdy += 'Bcc: ' + hdr['Bcc'] + '\n';
        for i in hdr['Bcc'].split(';'):
            ebcc.append(i);

    if 'Message-ID' in hdr:
        ebdy += 'Message-ID: ' + hdr['Message-ID'] + '\n';

    if 'Subject' in hdr:
         if hdr['Subject'] is not None:
             if len(hdr['Subject']) > 0 and isinstance(hdr['Subject'], str):
                 esbj += (decode_header(hdr['Subject'])[0])[0].decode('utf-8');
	
    if 'X-Asterisk-CallerID' in hdr and 'X-Asterisk-CallerIDName' in hdr:
        ebdy += 'Caller ID: ' + hdr['X-Asterisk-CallerID'] + ' (' + hdr['X-Asterisk-CallerIDName'] + ')\n';

    for b in bdy:
        if 'Content-Type' in bdy[b] and 'Content-Transfer-Encoding' in bdy[b] and 'Body' in bdy[b]:
            if re.match('text/plain;', bdy[b]['Content-Type']) and bdy[b]['Content-Transfer-Encoding'] == '8bit':
                ebdy += '\n\n';
                ebdy += "".join(bdy[b]['Body']);
                ebdy += '\n';

    if ebdy == '':
        ebdy = None;

    if esbj == '':
        ebdy = None;

    #pprint.pprint(hdr);
    #pprint.pprint(bdy);    
    #pprint.pprint(ews);

    ''' Step 0: Prerequisites '''
    if 'EWS_USER' not in ews or 'EWS_PASS' not in ews or 'EWS_SERVER' not in ews:
        _log('Failed to locate EWS credentials', 'CRIT', 'ERR');
        _exit(1);

    if ews['EWS_SERVER'] == 'AUTO':
        ews['EWS_SERVER'] = None;
        
    ''' Step 1: Initialize Office 365 Session '''

    sess = EWSSession(ews['EWS_SERVER'], ews['EWS_USER'], ews['EWS_PASS'], log_lvl);

    if sess.err:
        err.extend(sess.err)
        _exit(1);

    if sess.log:
        print("\n".join(sess.log));
    
    sess.err[:] = [];
    sess.log[:] = [];

    ''' Step 2: Create e-mail draft object '''
    email = EWSEmail(log_lvl);
    email.formatting('plain');
    if len(eto) > 0:
        email.recipients(eto);
    if esbj is not None:
        email.subject(esbj);
    if ebdy is not None:
        email.body(ebdy);
    if len(ecc) > 0:
        email.cc(ecc);
    if len(ebcc) > 0:
        email.bcc(ebcc);
    email.importance('High');
    email.mark_read('No');
    email.finalize();


    ''' Step 3: Submit e-mail draft to Office 365 '''

    sess.submit('draft', email.xml);

    if sess.err:
        err.extend(sess.err)
        _exit(1);

    if sess.log:
        print("\n".join(sess.log));

    sess.err[:] = [];
    sess.log[:] = [];


    ''' Step 4: Create e-mail attachments object '''

    attachment = EWSAttachment(sess.id, sess.changekey, log_lvl);
    for b in bdy:
        if 'Content-Type' in bdy[b] and 'Content-Transfer-Encoding' in bdy[b] and 'Body' in bdy[b]:
            if re.match('audio/x-wav;', bdy[b]['Content-Type']) and bdy[b]['Content-Transfer-Encoding'] == 'base64':
                bfn = None;
                bft = None
                m = re.match('(.*);\sname="(.*)"$', bdy[b]['Content-Type']);
                if m:
                    bft = m.group(1);
                    bfn = m.group(2);
                bbt = bytes("".join(bdy[b]['Body']), 'utf-8');
                attachment.add(bbt, bfn);
    attachment.finalize();

    if attachment.err:
        err.extend(attachment.err)
        _exit(1);

    if attachment.log:
        print("\n".join(attachment.log));

    attachment.err[:] = [];
    attachment.log[:] = [];


    ''' Step 5: Add attachments to the draft '''
    
    sess.submit('attachment', attachment.xml);

    if sess.err:
        err.extend(sess.err)
        _exit(1);

    if sess.log:
        print("\n".join(sess.log));

    sess.err[:] = [];
    sess.log[:] = [];


    ''' Step 6: Send the draft and move it to Sent Items folder '''

    sess.submit('send_and_save');

    if sess.err:
        err.extend(sess.err)
        _exit(1);

    if sess.log:
        print("\n".join(sess.log));

    sess.err[:] = [];
    sess.log[:] = [];

    _exit(0);


if __name__ == '__main__':
    main();
