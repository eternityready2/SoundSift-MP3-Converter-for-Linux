#!/bin/bash
set -e
export DEBIAN_FRONTEND=noninteractive
mkdir -p /repo/pool/main/x /repo/dists/stable/main/binary-amd64
cat > /repo/dists/stable/main/binary-amd64/Packages <<'PKG'
Package: demo404
Version: 1.0
Architecture: amd64
Maintainer: test <test@example.com>
Filename: pool/main/x/demo404_1.0_amd64.deb
Size: 1234
MD5sum: deadbeefdeadbeefdeadbeefdeadbeef
SHA256: deadbeefdeadbeefdeadbeefdeadbeefdeadbeefdeadbeefdeadbeefdeadbeef
Description: missing deb to trigger 404
PKG
apt-ftparchive release /repo/dists/stable > /repo/dists/stable/Release
python3 -m http.server 8000 -d /repo >/tmp/http.log 2>&1 &
server=$!
sleep 1
printf 'deb [trusted=yes] http://127.0.0.1:8000 stable main\n' > /etc/apt/sources.list.d/local404.list
apt-get update
apt-get install -y demo404
code=$?
kill $server || true
wait $server 2>/dev/null || true
exit $code
