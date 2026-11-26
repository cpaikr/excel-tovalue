#!/bin/sh
# Install the xlclean CLI on Linux x64 (glibc).
# Use this installer, the archive, and SHA256SUMS.txt from the same release.
set -eu
fail() { printf '%s\n' "$*" >&2; exit 1; }
[ "$#" -ge 2 ] && [ "$#" -le 3 ] || fail "Usage: sh install.sh ARCHIVE SHA256SUMS [BIN_DIRECTORY]"
archive=$1
checksums=$2
bin_dir=${3:-"$HOME/.local/bin"}
[ -f "$archive" ] && [ -f "$checksums" ] || fail "Archive and SHA256SUMS must be local files."
case "$(uname -s):$(uname -m)" in
  Linux:x86_64) ;;
  *) fail "No standalone archive is available for this operating system/architecture." ;;
esac
archive_name=$(basename "$archive")
case "$archive_name" in
  xlclean-cli-*-linux-x64.tar.gz) ;;
  *) fail "Expected an archive named xlclean-cli-<version>-linux-x64.tar.gz, not $archive_name." ;;
esac
version=${archive_name#xlclean-cli-}
version=${version%-linux-x64.tar.gz}
# Hash stdin: GNU sha256sum escapes its output when a file name has a backslash.
sha256() {
  if command -v sha256sum >/dev/null 2>&1; then
    sha256sum <"$1" | awk '{ print $1 }'
  else
    shasum -a 256 <"$1" | awk '{ print $1 }'
  fi
}
# SHA256SUMS.txt is written on Windows; tolerate a carriage return per line.
expected=$(tr -d '\r' <"$checksums" | awk -v name="$archive_name" '$2 == name { print $1 }')
[ "${#expected}" -eq 64 ] || fail "Missing or duplicate checksum for $archive_name."
case "$expected" in *[!0-9a-f]*) fail "Invalid SHA-256 checksum." ;; esac
[ "$(sha256 "$archive")" = "$expected" ] || fail "Archive checksum mismatch; the existing installation was not changed."
# Only the executable, its license terms, and the third-party notices, as
# regular files, may be extracted. Only the executable is installed.
[ "$(tar -tzf "$archive")" = "xlclean
LICENSE.md
THIRD-PARTY-NOTICES.txt" ] || fail "Unexpected archive contents."
[ "$(tar -tvzf "$archive" | cut -c1 | sort -u)" = '-' ] || fail "Archive must contain only regular files."
mkdir -p "$bin_dir"
bin_dir=$(CDPATH='' cd -- "$bin_dir" && pwd -P)
stage=$(mktemp -d "$bin_dir/.xlclean-install.XXXXXX")
trap 'rm -rf "$stage"' EXIT HUP INT TERM
tar -xzf "$archive" -C "$stage"
# The archive is built on Windows, which records no executable bit.
chmod 755 "$stage/xlclean"
# Check the candidate before a same-filesystem atomic rename replaces the old binary.
reported=$("$stage/xlclean" --version) || fail "The candidate failed to start; the existing installation was not changed."
[ "$reported" = "xlclean $version" ] || fail "The candidate reported '$reported' instead of 'xlclean $version'; the existing installation was not changed."
mv -f "$stage/xlclean" "$bin_dir/xlclean"
printf 'Installed xlclean %s to %s/xlclean\nAdd %s to PATH if needed.\n' "$version" "$bin_dir" "$bin_dir"
