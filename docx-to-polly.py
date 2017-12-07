#!/usr/bin/env python
import os
import sys
import getpass
import subprocess
from argparse import ArgumentParser
from collections import namedtuple
from contextlib import closing
from io import BytesIO



def docx_to_md(ns):
    base, ext = os.path.splitext(ns.file)
    ns.md = base + '.md'
    cmd = ['pandoc', '-o', ns.md, ns.file]
    subprocess.check_call(cmd)


def md_to_ssml(ns):
    with open(ns.md) as f:
        s = f.read()
    s = '<speak>' + s + '</speak>'
    s.replace('\n\n', '\n<break time="0.3s" />\n')
    s.replace('WAIT', '<break time="1.3s">')
    ns.ssml = s


def make_parser():
    p = ArgumentParser('docx-to-polly')
    p.add_argument('file')
    p.add_argument('-u', '--user', default=getpass.getuser())
    return p


def main(args=None):
    p = make_parser()
    ns = p.parse_args(arge=args)
    docx_to_md(ns)
    md_to_ssml(ns)
    print(ns.ssml)


if __name__ == '__main__':
    main()