#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
docker exec kara_db pg_dump -U postgres kara_taskdb \
  | gzip > "backups/kara_taskdb-$(date +%F_%H%M).sql.gz"
ls -1t backups/*.sql.gz | tail -n +15 | xargs -r rm --
