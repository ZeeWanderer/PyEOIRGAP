import argparse

from pz_1 import pz_1
from pz_2 import pz_2
from pz_3 import pz_3


def main():
    switcher = {"pz_1": pz_1,
                "pz_2": pz_2,
                "pz_3": pz_3,
                # "pz_4": pz_4,
                }

    parser = argparse.ArgumentParser("main.py")
    parser.add_argument('pz', choices=switcher.keys())
    parser.add_argument('preset', choices=['max', 'sergey'])

    args = parser.parse_args()

    preset = args.preset

    switcher.get(args.pz)(preset)


if __name__ == "__main__":
    main()
