#!/usr/bin/env python3
"""功能：离线重建合成语料并比较 ZIP 部件内容，不将 ZIP 时间戳差异当作布局变化；时间：2026-11-03；作者：lq。"""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / 'third_party/genoffice/fixtures'
manifest = json.loads((FIXTURES / 'manifest.json').read_text())


def verify_hash(relative, expected):
    """功能：独立校验已登记文件，源样本与生成器变更必须显式更新登记；时间：2026-11-03；作者：lq。"""
    actual = hashlib.sha256((FIXTURES / relative).read_bytes()).hexdigest()
    if actual != expected:
        raise ValueError(f'{relative}: SHA-256 不匹配')


def verify_source(relative, expected):
    """功能：集成注释之外的生成器内容必须与登记的上游源码哈希一致；时间：2026-11-03；作者：lq。"""
    body = (FIXTURES / relative).read_bytes().split(b'\n', 1)[1]
    if hashlib.sha256(body).hexdigest() != expected:
        raise ValueError(f'{relative}: 上游源码内容不匹配')


def payload(path):
    """功能：忽略 ZIP 容器日期，仅比较所有实际部件及内容；时间：2026-11-03；作者：lq。"""
    with zipfile.ZipFile(path) as package:
        # 功能：上游早期探针缺少未使用的 xmlns:r；标准规范化仅忽略命名空间声明形式，其他内容仍比较；时间：2026-11-03；作者：lq。
        return {name: ET.canonicalize(package.read(name).decode(), rewrite_prefixes=True)
                if name.endswith(('.xml', '.rels')) else package.read(name)
                for name in package.namelist() if not name.endswith('/')}


verify_hash(manifest['generatorHelper']['file'], manifest['generatorHelper']['sha256'])
verify_source(manifest['generatorHelper']['file'], manifest['generatorHelper']['upstreamSha256'])
rebuilt = 0
with tempfile.TemporaryDirectory(prefix='mini-office-fixtures-') as temporary:
    for fixture in manifest['fixtures']:
        verify_hash(fixture['file'], fixture['sha256'])
        provenance = fixture['provenance']
        if 'expectedText' in provenance:
            verify_hash(provenance['expectedText']['file'], provenance['expectedText']['sha256'])
        if 'generator' not in provenance:
            continue
        verify_hash(provenance['generator'], provenance['generatorSha256'])
        verify_source(provenance['generator'], provenance['upstreamGeneratorSha256'])
        generated = Path(temporary) / f"{fixture['id']}.docx"
        subprocess.run([sys.executable, str(FIXTURES / provenance['generator']), str(generated)],
                       cwd=temporary, check=True, env={**os.environ, 'PYTHONDONTWRITEBYTECODE': '1'})
        if payload(generated) != payload(FIXTURES / fixture['file']):
            raise ValueError(f"{fixture['id']}: 生成后的 OOXML 部件与固定样本不同")
        rebuilt += 1
print(json.dumps({'fixturesVerified': len(manifest['fixtures']), 'generatorsRebuilt': rebuilt,
                  'officeBaselineVerified': False}, ensure_ascii=False))
