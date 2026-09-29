#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)"
cd -- "$PROJECT_DIR"

if docker compose version >/dev/null 2>&1; then
    compose=(docker compose)
elif [[ -x "$PROJECT_DIR/test-output/docker-compose" ]]; then
    compose=("$PROJECT_DIR/test-output/docker-compose")
else
    echo '需要安装 Docker Compose 插件后再运行此脚本。' >&2
    exit 1
fi

docker info --format 'Docker Engine: {{.ServerVersion}}'
previous_image=""
if docker image inspect coc7-card:local >/dev/null 2>&1; then
    previous_image="coc7-card:backup-$(date +%Y%m%d-%H%M%S)"
    docker image tag coc7-card:local "$previous_image"
    echo "已保留旧镜像：$previous_image"
fi
docker build --progress=plain --tag coc7-card:local "$PROJECT_DIR"
if ! "${compose[@]}" --project-directory "$PROJECT_DIR" --file "$PROJECT_DIR/compose.yaml" \
    up --detach --no-build --wait --wait-timeout 180; then
    if [[ -n "$previous_image" ]]; then
        docker image tag "$previous_image" coc7-card:local
        "${compose[@]}" --project-directory "$PROJECT_DIR" --file "$PROJECT_DIR/compose.yaml" \
            up --detach --no-build --wait --wait-timeout 180
        echo "新版未能健康启动，已恢复镜像：$previous_image" >&2
    fi
    exit 1
fi
"${compose[@]}" --project-directory "$PROJECT_DIR" --file "$PROJECT_DIR/compose.yaml" ps
"${compose[@]}" --project-directory "$PROJECT_DIR" --file "$PROJECT_DIR/compose.yaml" \
    exec -T coc7 python -c "import urllib.request; print(urllib.request.urlopen('http://127.0.0.1:8765/api/health', timeout=5).read().decode())"
echo 'COC7_CONTAINER_START_COMPLETE'
