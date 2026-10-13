"""Local-only chart library and setlist editor. Runs with Python's built-in HTTP server."""
from __future__ import annotations
import argparse
import io
import json
import re
import threading
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit, unquote
import chart_engine as chart

ROOT=Path(__file__).resolve().parent
DATA=ROOT/'data'/'songs.json'
STATIC=ROOT/'static'
LOCK=threading.RLock()
MAX_BYTES=3_000_000

def load():
    with LOCK:
        try:return json.loads(DATA.read_text(encoding='utf-8'))
        except FileNotFoundError:return []

def save(songs):
    with LOCK:
        DATA.parent.mkdir(exist_ok=True)
        tmp=DATA.with_suffix('.tmp')
        tmp.write_text(json.dumps(songs,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        tmp.replace(DATA)

def clean_song(obj, fallback_id=None):
    if not isinstance(obj,dict):raise ValueError('Each song must be an object.')
    title=str(obj.get('title') or 'Untitled Song').strip()[:160]
    return dict(id=fallback_id or str(uuid.uuid4()),title=title,
                artist=str(obj.get('artist') or '').strip()[:160],
                content=str(obj.get('content') or '')[:40000],
                included=bool(obj.get('included',True)))

class Handler(BaseHTTPRequestHandler):
    def send(self,code,content,ctype='application/json',filename=None):
        if isinstance(content,str):content=content.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type',ctype)
        self.send_header('Content-Length',str(len(content)))
        self.send_header('Cache-Control','no-store')
        self.send_header('X-Content-Type-Options','nosniff')
        if filename:self.send_header('Content-Disposition',f'attachment; filename="{filename}"')
        self.end_headers();self.wfile.write(content)

    def out(self,code,payload):self.send(code,json.dumps(payload,ensure_ascii=False).encode('utf-8'))

    def input(self):
        n=int(self.headers.get('Content-Length','0'))
        if not 0<n<=MAX_BYTES:raise ValueError('Request body must be 1 byte to 3 MB.')
        return json.loads(self.rfile.read(n).decode('utf-8'))

    def do_GET(self):
        path=unquote(urlsplit(self.path).path)
        try:
            if path=='/api/songs':return self.out(200,load())
            if path=='/api/backup':
                return self.send(200,json.dumps(load(),indent=2,ensure_ascii=False).encode('utf-8'),
                                 'application/json; charset=utf-8','song-library-backup.json')
            if path.startswith('/api/preview/') and path.endswith('.png'):
                id=path[len('/api/preview/'):-4]
                song=next((s for s in load() if s['id']==id),None)
                if not song:return self.out(404,{'error':'Song not found.'})
                return self.send(200,chart.preview(song),'image/png')
            if path in ('/api/export/pptx','/api/export/pdf'):
                songs=[s for s in load() if s.get('included',True)]
                if not songs:return self.out(422,{'error':'Select at least one song to export.'})
                if path.endswith('/pptx'):
                    return self.send(200,chart.pptx(songs),
                              'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                              'live-show-charts.pptx')
                return self.send(200,chart.pdf(songs),'application/pdf','live-show-charts.pdf')
            resources={'/':'index.html','/index.html':'index.html',
                       '/app.js':'app.js','/styles.css':'styles.css'}
            if path in resources:
                ext=path.rsplit('.',1)[-1]
                ctype='text/html; charset=utf-8' if ext=='html' or path=='/' else (
                    'text/javascript; charset=utf-8' if ext=='js' else 'text/css; charset=utf-8')
                return self.send(200,(STATIC/resources[path]).read_bytes(),ctype)
            return self.out(404,{'error':'Not found.'})
        except chart.ChartOverflow as e:return self.out(422,{'error':str(e)})
        except Exception as e:
            print('GET error:',repr(e));return self.out(500,{'error':str(e)})

    def do_POST(self):
        path=urlsplit(self.path).path
        try:
            payload=self.input()
            if path=='/api/songs':
                song=clean_song(payload)
                with LOCK:
                    songs=load();songs.append(song);save(songs)
                return self.out(201,song)
            if path=='/api/order':
                ids=payload.get('ids') if isinstance(payload,dict) else None
                with LOCK:
                    songs=load()
                    if not isinstance(ids,list) or len(ids)!=len(songs) or set(ids)!={s['id'] for s in songs}:
                        raise ValueError('Order must contain each song exactly once.')
                    index={s['id']:s for s in songs}
                    songs=[index[i] for i in ids]
                    save(songs)
                return self.out(200,{'ok':True})
            if path=='/api/restore':
                if not isinstance(payload,list) or len(payload)>500:
                    raise ValueError('Import a JSON array with no more than 500 songs.')
                songs=[];seen=set()
                for raw in payload:
                    song=clean_song(raw)
                    old=raw.get('id') if isinstance(raw,dict) else None
                    if isinstance(old,str) and re.fullmatch('[A-Za-z0-9_-]{1,100}',old) and old not in seen:
                        song['id']=old
                    seen.add(song['id']);songs.append(song)
                save(songs)
                return self.out(200,{'ok':True,'count':len(songs)})
            return self.out(404,{'error':'Not found.'})
        except (ValueError,TypeError,UnicodeDecodeError,json.JSONDecodeError) as e:return self.out(400,{'error':str(e)})
        except Exception as e:
            print('POST error:',repr(e));return self.out(500,{'error':str(e)})

    def do_PUT(self):
        path=urlsplit(self.path).path
        try:
            obj=self.input()
            if not path.startswith('/api/song/') or not isinstance(obj,dict):
                return self.out(404,{'error':'Not found.'})
            id=path[len('/api/song/'):]
            with LOCK:
                songs=load()
                for i,s in enumerate(songs):
                    if s['id']==id:
                        songs[i]=clean_song({**s,**obj},fallback_id=id)
                        save(songs)
                        return self.out(200,songs[i])
            return self.out(404,{'error':'Song not found.'})
        except (ValueError,TypeError) as e:return self.out(400,{'error':str(e)})
        except Exception as e:
            print('PUT error:',repr(e));return self.out(500,{'error':str(e)})

    def do_DELETE(self):
        path=urlsplit(self.path).path
        if not path.startswith('/api/song/'):return self.out(404,{'error':'Not found.'})
        id=path[len('/api/song/'):]
        with LOCK:
            songs=load();remaining=[s for s in songs if s['id']!=id]
            if len(remaining)==len(songs):return self.out(404,{'error':'Song not found.'})
            save(remaining)
        return self.out(200,{'ok':True})

    def log_message(self,format,*args):
        if not self.path.startswith('/api/preview/'):
            super().log_message(format,*args)

if __name__=='__main__':
    parser=argparse.ArgumentParser(description='Local live-show song chart generator')
    parser.add_argument('--port',type=int,default=8765)
    parser.add_argument('--host',default='127.0.0.1',help='Bind address (local-only by default)')
    parser.add_argument('--no-browser',action='store_true')
    args=parser.parse_args()
    url=f'http://localhost:{args.port}'
    server=ThreadingHTTPServer((args.host,args.port),Handler)
    print(f'\nSong Chart Generator running at {url}\nPress Ctrl+C to stop.\n')
    if not args.no_browser:threading.Timer(0.7,lambda:webbrowser.open(url)).start()
    try:server.serve_forever()
    except KeyboardInterrupt:pass
    finally:server.server_close()
