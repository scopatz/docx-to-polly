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
    s = '<speak>' + s + '</speak>'
    ns.ssml = s


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
    print(ns.ssml)


if __name__ == '__main__':
    main()