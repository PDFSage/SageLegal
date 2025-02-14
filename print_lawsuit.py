import pickle
import os

def load_dictionary_from_pickle(filepath):
    """
    Load a dictionary from a pickle file.

    Args:
        filepath (str): The path to the pickle file.

    Returns:
        dict: The dictionary loaded from the pickle file.

    Raises:
        FileNotFoundError: If the pickle file does not exist.
        ValueError: If the loaded data is not a dictionary.
        Exception: For any other errors during the unpickling process.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"The file '{filepath}' does not exist.")

    try:
        with open(filepath, 'rb') as file:
            data = pickle.load(file)
    except Exception as e:
        raise Exception(f"Error unpickling file '{filepath}': {e}") from e

    if not isinstance(data, dict):
        raise ValueError(f"Expected a dictionary, but got {type(data).__name__}.")

    return data

def main():
    pickle_file = 'lawsuit.pickle'
    
    try:
        lawsuit_obj = load_dictionary_from_pickle(pickle_file)
        print("Loaded dictionary from pickle:")
        print(lawsuit_obj)
    except Exception as error:
        print(f"An error occurred: {error}")

if __name__ == '__main__':
    main()