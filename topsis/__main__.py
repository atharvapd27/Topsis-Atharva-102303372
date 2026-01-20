from .topsis import topsis
import sys

if __name__ == "__main__":
    if len(sys.argv) != 5:
        print("Usage: python -m topsis <InputFile> <Weights> <Impacts> <OutputFile>")
        sys.exit(1)

    topsis(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
