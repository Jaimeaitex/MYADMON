"""Asistente para redactar partes usando GitHub Copilot CLI (modo no interactivo, sin herramientas)."""
from __future__ import annotations
import os
from pathlib import Path
import re
import shutil
import subprocess

INSTALL_HELP = ('No se encuentra GitHub Copilot CLI.\n'
                'Instálalo con:  npm install -g @github/copilot\n'
                'y ejecuta «copilot» una vez en una terminal para iniciar sesión.')

SYSTEM = """Eres un asistente que ayuda a un técnico de AITEX (centro tecnológico textil) a redactar sus partes horarios.
Cada parte es una línea breve en español que describe el trabajo realizado en un proyecto, con el estilo de los partes anteriores.
Reglas:
- Propón actividades coherentes con la evolución reciente del proyecto (siguiente paso lógico, continuidad de tareas).
- Imita el tono y la longitud de los partes anteriores. No inventes resultados, cifras ni nombres que no aparezcan en el contexto.
- Responde siempre en español y de forma breve.
- Cuando propongas un parte concreto, termina con estas dos líneas exactas:
PARTE: <descripción lista para copiar>
HORAS: <H:MM sugeridas, opcional>"""


def _command():
    node = shutil.which('node')
    roots = [Path(os.environ.get('APPDATA', '')) / 'npm' / 'node_modules']
    try:
        out = subprocess.run(['npm', 'root', '-g'], capture_output=True, text=True, timeout=15,
                             shell=os.name == 'nt', creationflags=_NO_WINDOW)
        if out.stdout.strip():
            roots.append(Path(out.stdout.strip()))
    except (OSError, subprocess.SubprocessError):
        pass
    for root in roots:
        loader = root / '@github' / 'copilot' / 'npm-loader.js'
        if node and loader.is_file():
            return [node, str(loader)]
    return None


_NO_WINDOW = 0x08000000 if os.name == 'nt' else 0
_cached = []


def copilot_command():
    if not _cached:
        _cached.append(_command())
    return _cached[0]


def build_prompt(context, history, question):
    lines = [SYSTEM, '', '=== CONTEXTO ===', context, '', '=== CONVERSACIÓN ===']
    for role, text in history:
        lines.append(f'{"Usuario" if role == "user" else "Asistente"}: {text}')
    lines.append(f'Usuario: {question}')
    lines.append('Asistente:')
    return '\n'.join(lines)


def ask(prompt, workdir, timeout=180, on_start=None):
    """Envía el prompt a Copilot y devuelve la respuesta en texto."""
    cmd = copilot_command()
    if not cmd:
        raise RuntimeError(INSTALL_HELP)
    Path(workdir).mkdir(parents=True, exist_ok=True)
    proc = subprocess.Popen(cmd + ['-s', '--no-custom-instructions', '--available-tools'],
                            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            cwd=workdir, creationflags=_NO_WINDOW)
    if on_start:
        on_start(proc)
    try:
        out, err = proc.communicate(prompt.encode('utf-8'), timeout=timeout)
    except subprocess.TimeoutExpired:
        proc.kill()
        raise RuntimeError('Copilot ha tardado demasiado en responder. Inténtalo de nuevo.')
    text = out.decode('utf-8', 'replace').strip()
    if proc.returncode != 0 or not text:
        detail = err.decode('utf-8', 'replace').strip() or text
        if re.search(r'auth|login|token|sign', detail, re.I):
            detail += '\n\nEjecuta «copilot» en una terminal e inicia sesión con tu cuenta de GitHub.'
        raise RuntimeError(detail or f'Copilot terminó con código {proc.returncode}')
    return text


def parse_suggestion(text):
    """Extrae (descripción, minutos) de las líneas PARTE:/HORAS: de la respuesta."""
    desc = None
    minutes = None
    for line in text.splitlines():
        m = re.match(r'\s*\**PARTE\**\s*:\s*\**(.+?)\**\s*$', line, re.I)
        if m:
            desc = m.group(1).strip().strip('"«»')
        m = re.match(r'\s*\**HORAS\**\s*:\s*\**\s*(\d{1,2})(?:\s*([:.,h])\s*(\d{1,2}))?', line, re.I)
        if m:
            h, sep, frac = int(m.group(1)), m.group(2), m.group(3) or '0'
            mm = int(frac) * 6 if sep in ('.', ',') and len(frac) == 1 else int(frac)
            if mm < 60:
                minutes = h * 60 + mm
    return desc, minutes
