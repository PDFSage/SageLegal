#!/usr/bin/env python3

import pickle

def main():
    pickle_filename = "lawsuit.pickle"  # Adjust if you saved to a different filename

    with open(pickle_filename, 'rb') as pf:
        loaded_obj = pickle.load(pf)
        print("Loaded Lawsuit object from pickle:")
        print(loaded_obj)

if __name__ == "__main__":
    main()