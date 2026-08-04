import urllib.request
url_openai = 'https://raw.githubusercontent.com/simple-icons/simple-icons/master/icons/openai.svg'

try:
    req = urllib.request.Request(url_openai)
    with urllib.request.urlopen(req) as response:
        content = response.read().decode('utf-8')
        content = content.replace('<svg ', '<svg fill="#10A37F" ') # Inject color
        with open('images/openai.svg', 'w', encoding='utf-8') as f:
            f.write(content)
    print('OpenAI master saved.')
except Exception as e:
    print('OpenAI failed:', e)

url_claude = 'https://raw.githubusercontent.com/simple-icons/simple-icons/master/icons/claude.svg'
try:
    req = urllib.request.Request(url_claude)
    with urllib.request.urlopen(req) as response:
        content = response.read().decode('utf-8')
        content = content.replace('<svg ', '<svg fill="#D97757" ') # Inject color
        with open('images/claude.svg', 'w', encoding='utf-8') as f:
            f.write(content)
    print('Claude master saved.')
except Exception as e:
    print('Claude failed:', e)
