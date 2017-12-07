#!/usr/bin/env python
import os
import sys
import re
import getpass
import subprocess
from argparse import ArgumentParser
from collections import namedtuple
from contextlib import closing
from io import BytesIO

from boto3 import Session
from botocore.exceptions import BotoCoreError, ClientError


RE_UNDERLINE = re.compile('\[(.*?)\]\{\.underline\}')


def docx_to_md(ns):
    base, ext = os.path.splitext(ns.file)
    ns.md = base + '.md'
    cmd = ['pandoc', '-o', ns.md, ns.file]
    subprocess.check_call(cmd)


def md_to_ssmls(ns):
    with open(ns.md) as f:
        s = f.read()
    s = s.replace('WAIT', '<break time="1.3s" />')
    s = RE_UNDERLINE.sub(
            lambda m: '<emphasis level="strong">' + m.group(1) + '</emphasis>', s)
    lines = []
    for line in s.splitlines():
        if line.startswith('=') and len(set(line)) == 1:
            if len(lines) > 1:
                lines.insert(-2, '<break time="0.3s"/>')
        else:
            lines.append(line)
    s = '\n'.join(lines)
    # brak up by paragraph
    blocks = s.split('\n\n')
    for i in range(len(blocks)):
        blocks[i] = '<speak>' + blocks[i] + '<break time="0.3s" /></speak>'
    ns.ssmls = blocks


def session(ns):
    ns.session = Session(profile_name=ns.user)
    ns.polly = ns.session.client("polly")


def tts(ns):
    nblocks = len(ns.ssmls)
    aud = b''
    for i, ssml in enumerate(ns.ssmls):
        print('{0}/{1}\r'.format(i, nblocks), end='')
        sys.stdout.flush()
        response = ns.polly.synthesize_speech(Text=ssml,
                                              TextType='ssml',
                                              VoiceId='Matthew',
                                              OutputFormat='mp3')
        aud += response["AudioStream"].read()
    base, ext = os.path.splitext(ns.file)
    ns.out = base + '.mp3'
    with open(ns.out, 'bw') as f:
        f.write(aud)
    print('Done... ' + ns.out)


def make_parser():
    p = ArgumentParser('docx-to-polly')
    p.add_argument('file')
    p.add_argument('-u', '--user', default=getpass.getuser())
    return p


def main(args=None):
    p = make_parser()
    ns = p.parse_args(args=args)
    docx_to_md(ns)
    md_to_ssmls(ns)
    session(ns)
    tts(ns)

if __name__ == '__main__':
    main()