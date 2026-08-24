out_file = r'C:\Users\congn\Pictures\vnptemploy\blutter_api_analysis.txt'

with open(r'C:\Users\congn\Pictures\vnptemploy\blutter_out\pp.txt', encoding='utf-8', errors='replace') as f:
    content = f.read()

lines = content.splitlines()

# Keywords liên quan API/HTTP/Login
api_kws = ['http', 'login', 'msisdn', 'password', 'token', 'auth', 'otp',
           '/api/', 'request', 'response', 'header', 'bearer', 'dio',
           'retrofit', 'interceptor', 'baseUrl', 'baseurl', 'onebss',
           'ekyc', 'idg', 'videocall', 'service', 'endpoint', 'client',
           'post(', 'get(', 'put(', 'delete(']

results = []
i = 0
while i < len(lines):
    line = lines[i]
    if any(k in line.lower() for k in api_kws):
        # Lấy cả block (class/method context)
        start = max(0, i - 3)
        end = min(len(lines), i + 5)
        block = '\n'.join(lines[start:end])
        results.append(f'[L{i+1}]\n{block}')
        i = end  # skip đã include
    else:
        i += 1

print(f'Total relevant blocks: {len(results)}')

with open(out_file, 'w', encoding='utf-8') as f:
    f.write(f'=== BLUTTER - VNPT Employee Dart API Analysis ===\n')
    f.write(f'Total relevant blocks: {len(results)}\n')
    f.write('=' * 70 + '\n\n')
    for r in results:
        f.write(r + '\n\n')

print(f'Saved to: {out_file}')
print('\n--- PREVIEW (first 80 blocks) ---')
for r in results[:80]:
    print(r)
    print()
