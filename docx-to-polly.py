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


def md_to_ssml(ns):
    with open(ns.md) as f:
        s = f.read()
    s = s.replace('\n\n', '\n<break time="0.3s" />\n\n')
    s = s.replace('WAIT', '<break time="1.3s">')
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
    s = '<speak>' + s[:100] + '</speak>'
    ns.ssml = s


def session(ns):
    ns.session = Session(profile_name=ns.user)
    ns.polly = ns.session.client("polly")


def tts(ns):
    response = ns.polly.synthesize_speech(Text=ns.ssml,
                                          TextType='ssml',
                                          VoiceId='Matthew',
                                          OutputFormat='mp3')
    aud = response["AudioStream"].read()
    base, ext = os.path.splitext(ns.file)
    ns.out = base + '.mp3'
    with open(ns.out, 'bw') as f:
        f.write(aud)


def make_parser():
    p = ArgumentParser('docx-to-polly')
    p.add_argument('file')
    p.add_argument('-u', '--user', default=getpass.getuser())
    return p


def main(args=None):
    p = make_parser()
    ns = p.parse_args(args=args)
    docx_to_md(ns)
    md_to_ssml(ns)
    session(ns)
    tts(ns)

if __name__ == '__main__':
    main()