import os
import zipfile


def read_manifest(path):
    data = {}
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or "=" not in line:
                continue
            key, value = line.split("=", 1)
            value = value.strip().strip('"')
            data[key.strip()] = value
    return data


def add_dir(zf, base):
    for root, _dirs, files in os.walk(base):
        for name in files:
            full_path = os.path.join(root, name)
            rel_path = full_path.replace("\\", "/")
            zf.write(full_path, rel_path)


if __name__ == "__main__":
    manifest = "manifest.ini"
    info = read_manifest(manifest)
    name = info.get("name", "addon")
    version = info.get("version", "0.0.0")

    out_dir = "dist"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"{name}-{version}.nvda-addon")
    if os.path.exists(out_path):
        os.remove(out_path)

    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(manifest, "manifest.ini")
        if os.path.isfile("LICENSE"):
            zf.write("LICENSE", "LICENSE")
        if os.path.isdir("doc"):
            add_dir(zf, "doc")
        if os.path.isdir("globalPlugins"):
            add_dir(zf, "globalPlugins")

    print(out_path)
