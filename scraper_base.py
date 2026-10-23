# -*- coding: utf-8 -*-
"""
scraper_base.py — base reutilizável para raspagem educada de páginas HTML.

O que tem aqui:
  - criar_sessao(): sessão requests com User-Agent identificado
  - baixar(url): download com 3 tentativas, backoff exponencial e delay entre
    requisições (para não sobrecarregar o site de origem)
  - pode_raspar(url): consulta o robots.txt do site antes de baixar
  - parsear / buscar / texto: um parser minimalista em cima do html.parser
    da biblioteca padrão (sem BeautifulSoup), suficiente para localizar
    elementos por tag e classe CSS

Autor: Marcos Galani — automação de processos
"""
from __future__ import annotations

import logging
import time
from html.parser import HTMLParser
from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser

import requests

log = logging.getLogger("monitor")

# ---------------------------------------------------------------------------
# Configuração de "boa vizinhança"
# ---------------------------------------------------------------------------
USER_AGENT = (
    "monitor-precos-demo/1.0 (demo de portfolio; scraping educado; "
    "contato via plataforma freelance)"
)
DELAY_ENTRE_PAGINAS = 1.0   # segundos de pausa entre downloads
TENTATIVAS = 3              # quantas vezes tentar antes de desistir
BACKOFF_BASE = 2.0          # espera 2 s, depois 4 s, depois 8 s...
TIMEOUT = 20                # segundos por requisição

_ultimo_download = 0.0      # controle interno do delay
_robots_cache: dict[str, RobotFileParser] = {}


def criar_sessao() -> requests.Session:
    """Sessão HTTP com cabeçalhos identificados (o site sabe quem somos)."""
    sessao = requests.Session()
    sessao.headers.update({
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml",
        "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
    })
    return sessao


def pode_raspar(url: str, sessao: requests.Session | None = None) -> bool:
    """
    Verifica o robots.txt do domínio. Se o arquivo não existir ou não puder
    ser lido, assume liberado (comportamento padrão do robotparser), mas
    registra no log para o operador conferir os termos de uso manualmente.
    """
    partes = urlparse(url)
    raiz = f"{partes.scheme}://{partes.netloc}"
    if raiz not in _robots_cache:
        rp = RobotFileParser()
        try:
            s = sessao or criar_sessao()
            resp = s.get(urljoin(raiz, "/robots.txt"), timeout=TIMEOUT)
            if resp.status_code == 200:
                rp.parse(resp.text.splitlines())
                log.info("robots.txt lido em %s", raiz)
            else:
                rp.allow_all = True
                log.info("robots.txt não encontrado em %s (HTTP %s); "
                         "confira os termos de uso do site", raiz, resp.status_code)
        except requests.RequestException as e:
            rp.allow_all = True
            log.warning("não consegui ler robots.txt de %s: %s", raiz, e)
        _robots_cache[raiz] = rp
    return _robots_cache[raiz].can_fetch(USER_AGENT, url)


def baixar(url: str, sessao: requests.Session | None = None) -> str:
    """
    Baixa uma URL e devolve o HTML como texto (UTF-8).

    - respeita DELAY_ENTRE_PAGINAS entre downloads consecutivos
    - tenta até TENTATIVAS vezes com backoff exponencial em erro de rede,
      HTTP 5xx ou 429
    - levanta exceção se todas as tentativas falharem
    """
    global _ultimo_download
    s = sessao or criar_sessao()

    if not pode_raspar(url, s):
        raise PermissionError(f"robots.txt proíbe acessar {url}")

    ultimo_erro: Exception | None = None
    for tentativa in range(1, TENTATIVAS + 1):
        # delay educado: garante pelo menos DELAY_ENTRE_PAGINAS desde o último GET
        espera = DELAY_ENTRE_PAGINAS - (time.monotonic() - _ultimo_download)
        if espera > 0:
            time.sleep(espera)
        try:
            resp = s.get(url, timeout=TIMEOUT)
            _ultimo_download = time.monotonic()
            if resp.status_code in (429, 500, 502, 503, 504):
                raise requests.HTTPError(f"HTTP {resp.status_code}", response=resp)
            resp.raise_for_status()
            # a página declara UTF-8 no <meta>; não confiar no palpite do requests
            return resp.content.decode("utf-8", errors="replace")
        except requests.RequestException as e:
            ultimo_erro = e
            if tentativa < TENTATIVAS:
                pausa = BACKOFF_BASE ** tentativa
                log.warning("tentativa %d/%d falhou em %s (%s); nova tentativa em %.0f s",
                            tentativa, TENTATIVAS, url, e, pausa)
                time.sleep(pausa)
    raise RuntimeError(f"falha ao baixar {url} após {TENTATIVAS} tentativas: {ultimo_erro}")


# ---------------------------------------------------------------------------
# Parser HTML minimalista (stdlib) — árvore de nós + busca por tag/classe
# ---------------------------------------------------------------------------
_VAZIAS = {"img", "br", "hr", "input", "meta", "link", "source", "wbr", "area", "base", "col"}


class No:
    """Nó de uma árvore HTML simplificada."""
    __slots__ = ("tag", "attrs", "filhos", "pai")

    def __init__(self, tag: str, attrs: dict, pai: "No | None"):
        self.tag = tag
        self.attrs = attrs
        self.filhos: list = []   # strings (texto) ou outros No
        self.pai = pai

    def classes(self) -> list[str]:
        return (self.attrs.get("class") or "").split()

    def tem_classe(self, classe: str) -> bool:
        return classe in self.classes()

    def __repr__(self) -> str:  # ajuda no debug
        return f"<{self.tag} {self.attrs}>"


class ConstruirDOM(HTMLParser):
    """Transforma o HTML numa árvore de No a partir dos eventos do html.parser."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.raiz = No("#raiz", {}, None)
        self._atual = self.raiz

    def handle_starttag(self, tag, attrs):
        no = No(tag, dict(attrs), self._atual)
        self._atual.filhos.append(no)
        if tag not in _VAZIAS:
            self._atual = no

    def handle_endtag(self, tag):
        if tag in _VAZIAS:
            return
        # sobe até o nó de mesma tag (tolera HTML mal fechado)
        no = self._atual
        while no is not None and no.tag != tag:
            no = no.pai
        if no is not None and no.pai is not None:
            self._atual = no.pai

    def handle_data(self, data):
        if data:
            self._atual.filhos.append(data)


def parsear(html: str) -> No:
    """HTML (texto) -> árvore de nós."""
    p = ConstruirDOM()
    p.feed(html)
    p.close()
    return p.raiz


def buscar(no: No, tag: str | None = None, classe: str | None = None) -> list[No]:
    """
    Busca descendentes que casem com a tag e/ou classe CSS.
    Equivale a um `find_all(tag, class_=classe)` simplificado.
    """
    achados: list[No] = []
    for filho in no.filhos:
        if isinstance(filho, No):
            if (tag is None or filho.tag == tag) and (classe is None or filho.tem_classe(classe)):
                achados.append(filho)
            achados.extend(buscar(filho, tag, classe))
    return achados


def primeiro(no: No, tag: str | None = None, classe: str | None = None) -> No | None:
    r = buscar(no, tag, classe)
    return r[0] if r else None


def texto(no: No | None) -> str:
    """Texto visível do nó e dos descendentes, com espaços normalizados."""
    if no is None:
        return ""
    partes: list[str] = []

    def _coleta(n: No):
        for f in n.filhos:
            if isinstance(f, str):
                partes.append(f)
            else:
                _coleta(f)

    _coleta(no)
    return " ".join("".join(partes).split())
