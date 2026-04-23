#!/usr/bin/env python3
"""Teste automatizado de busca de respostas do Apps Script.

Uso:
  python test_fetch_responses.py --script-url "https://script.google.com/macros/s/XXX/exec"

O script tenta obter:
  - respostas (?tipo=respostas)
  - links (?tipo=links)

Ele valida que o servidor responde com JSON ou JSONP e grava arquivos
na pasta do projeto:
  - responses.json
  - links.json

Se o Apps Script não estiver configurado corretamente, o teste exibirá
erros mais claros sobre o retorno.
"""

import argparse
import json
import os
import re
import sys
from urllib.parse import urlencode, urlparse, parse_qsl, urlunparse
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

JSONP_PATTERN = re.compile(r"^([a-zA-Z_$][0-9a-zA-Z_$]*)\((.*)\);?$", re.S)


def build_url(base_url, params):
    parsed = urlparse(base_url)
    query = dict(parse_qsl(parsed.query))
    query.update(params)
    new_query = urlencode(query)
    return urlunparse(parsed._replace(query=new_query))


def fetch_url(url):
    req = Request(url, headers={
        'User-Agent': 'AHSD-Test/1.0',
        'Accept': '*/*',
    })
    with urlopen(req, timeout=20) as resp:
        raw = resp.read().decode('utf-8', errors='replace')
        content_type = resp.headers.get('Content-Type', '')
    return raw, content_type


def parse_response(raw):
    raw = raw.strip()
    if not raw:
        raise ValueError('Resposta vazia do servidor')

    m = JSONP_PATTERN.match(raw)
    if m:
        raw = m.group(2).strip()

    return json.loads(raw)


def save_json(data, path):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def run_test(script_url, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    results = {}

    for tipo in ('respostas', 'links'):
        print(f'--- Testando endpoint: tipo={tipo}')
        for mode in ('json', 'jsonp'):
            try:
                if mode == 'json':
                    url = build_url(script_url, {'tipo': tipo})
                else:
                    url = build_url(script_url, {'tipo': tipo, 'callback': 'testCb'})
                print(f'  solicitando {mode.upper()} -> {url}')
                raw, content_type = fetch_url(url)
                data = parse_response(raw)
                print(f'    OK: recebeu {len(data) if isinstance(data, list) else "objeto"}')
                results[(tipo, mode)] = {
                    'url': url,
                    'content_type': content_type,
                    'data': data,
                }
                break
            except (HTTPError, URLError) as exc:
                print(f'    erro de rede: {exc}')
                last_exc = exc
            except ValueError as exc:
                print(f'    parse failed: {exc}')
                print('    resposta bruta:', raw[:1200].replace('\n', ' '))
                last_exc = exc
            except json.JSONDecodeError as exc:
                print(f'    JSON decode failed: {exc}')
                print('    resposta bruta:', raw[:1200].replace('\n', ' '))
                last_exc = exc
        else:
            raise RuntimeError(f'Não foi possível obter {tipo}: {last_exc}')

        basename = f'{tipo}.json'
        path = os.path.join(output_dir, basename)
        save_json(results[(tipo, mode)]['data'], path)
        print(f'    gravado em: {path}')

    print('\nTeste concluído com sucesso.')
    return results


def parse_args():
    parser = argparse.ArgumentParser(description='Teste de fetch de respostas do Apps Script AH/SD')
    parser.add_argument('--script-url', required=True, help='URL de implantação do Apps Script (deve terminar em /exec)')
    parser.add_argument('--output-dir', default='.', help='Diretório de saída para os arquivos JSON')
    return parser.parse_args()


if __name__ == '__main__':
    args = parse_args()
    if not args.script_url.strip():
        print('Informe a URL do Apps Script com --script-url')
        sys.exit(1)

    try:
        run_test(args.script_url.strip(), args.output_dir)
    except Exception as exc:
        print(f'ERRO: {exc}')
        sys.exit(1)
