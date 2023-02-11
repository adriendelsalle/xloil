from pathlib import Path


for p in Path().iterdir():
    if p.is_dir():
        for f in p.iterdir():
            if f.suffix == ".cpp":
                print("${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/" + str(f))