#!/bin/bash
# Runs one mode over one folder; if the process dies (out of memory), records the document it died on and carries on.
cd $(git rev-parse --show-toplevel); mode=$1; folder=$2; out=$3; log=$4
total=$(ls $folder | grep -ci '\.docx$')
while [ "$(wc -l < $out 2>/dev/null || echo 0)" -lt "$total" ]; do
  MODE=$mode PDF_URL=http://127.0.0.1:8099/convert npx tsx scripts/corpus/run.ts $folder $out >> $log 2>&1
  [ "$(wc -l < $out)" -ge "$total" ] && break
  python3 - "$folder" "$out" <<'PY'
import json,os,sys
folder,out=sys.argv[1:]
done={json.loads(l)["file"] for l in open(out) if l.strip()}
for f in sorted(n for n in os.listdir(folder) if n.lower().endswith(".docx")):
    if f not in done:
        open(out,"a").write(json.dumps({"file":f,"error":"process stopped: out of memory or crash"})+"\n"); break
PY
done
echo "$mode $folder finished: $(wc -l < $out)"
