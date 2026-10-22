"""統合テスト用：ライブ分割を別プロセスで1回動かす（強制終了の試験で使う）。

python live_worker.py <設定JSON> <home> <プロファイル名> <pst_key>
"""

import json
import sys
from pathlib import Path

from pstmigration import config, pipeline

cfg_path, home, profile, key = sys.argv[1:]
cfg = config.Config(**json.loads(Path(cfg_path).read_text(encoding="utf-8")))
r = pipeline.Runner(cfg, Path(home))
r.profile = profile
r._cleanup_temp_profiles()  # 本番の起動時（Runner.run）と同じく、前回の残りを先に消す
e = r.st.psts[key]
r._process_live(e)
r.st.save()
print(json.dumps({"status": e["status"], "last_error": e.get("last_error")}, ensure_ascii=False))
