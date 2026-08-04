import urllib.request
import json

# Let's search for SVG content from simple-icons for Claude and OpenAI, maybe the previous download failed because I used the wrong URL.
url_claude = 'https://raw.githubusercontent.com/simple-icons/simple-icons/develop/icons/claude.svg'
url_openai = 'https://raw.githubusercontent.com/simple-icons/simple-icons/develop/icons/openai.svg'

try:
    req = urllib.request.Request(url_claude)
    with urllib.request.urlopen(req) as response:
        with open('images/claude.svg', 'wb') as f:
            f.write(response.read())
    print('Claude saved.')
except Exception as e:
    print('Claude failed:', e)

try:
    req = urllib.request.Request(url_openai)
    with urllib.request.urlopen(req) as response:
        with open('images/openai.svg', 'wb') as f:
            f.write(response.read())
    print('OpenAI saved.')
except Exception as e:
    print('OpenAI failed:', e)
