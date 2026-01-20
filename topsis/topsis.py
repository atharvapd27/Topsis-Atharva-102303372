import sys
import pandas as pd
import numpy as np
import os

def topsis(input_file, weights, impacts, output_file):
    #reading input file csv or excel
    try:
        data = pd.read_excel(input_file)
    #handling of “File not Found” exception
    except FileNotFoundError:
        print("Error: Input file not found.")
        sys.exit(1)

    #handling of the input file must contain three or more columns exception
    if data.shape[1] < 3:
        print("Error: Input file must contain at least three columns.")
        sys.exit(1)

    #handling of the criteria columns must contain numeric values only exception
    criteria = data.iloc[:, 1:]
    try:
        criteria = criteria.astype(float)
    except ValueError:
        print("Error: All criteria values must be numeric.")
        sys.exit(1)

    #handling of impacts and weights must be separated by commas exception
    weights = weights.split(',')
    impacts = impacts.split(',')

    #handling of number of weights and impacts must be equal to number of criteria exception
    if len(weights) != criteria.shape[1]:
        print("Error: Number of weights must match number of criteria.")
        sys.exit(1)

    if len(impacts) != criteria.shape[1]:
        print("Error: Number of impacts must match number of criteria.")
        sys.exit(1)

    #handling of impacts must be either '+' or '-' exception
    if not all(i in ['+', '-'] for i in impacts):
        print("Error: Impacts must be either '+' or '-'.")
        sys.exit(1)

    weights = np.array(weights, dtype=float)

    # Step 1: Normalize
    norm = criteria / np.sqrt((criteria ** 2).sum())

    # Step 2: Weighted normalization
    weighted = norm * weights

    # Step 3: Ideal best & worst
    ideal_best = []
    ideal_worst = []

    for i in range(len(impacts)):
        if impacts[i] == '+':
            ideal_best.append(weighted.iloc[:, i].max())
            ideal_worst.append(weighted.iloc[:, i].min())
        else:
            ideal_best.append(weighted.iloc[:, i].min())
            ideal_worst.append(weighted.iloc[:, i].max())

    ideal_best = np.array(ideal_best)
    ideal_worst = np.array(ideal_worst)

    # Step 4: Distance calculation
    dist_best = np.sqrt(((weighted - ideal_best) ** 2).sum(axis=1))
    dist_worst = np.sqrt(((weighted - ideal_worst) ** 2).sum(axis=1))

    # Step 5: Performance score
    score = dist_worst / (dist_best + dist_worst)

    data['Topsis Score'] = score
    data['Rank'] = data['Topsis Score'].rank(ascending=False).astype(int)

    #save output file
    data.to_excel(output_file, index=False)
    print("Result saved to", output_file)


if __name__ == "__main__":
    #handling of incorrect number of parameters exception
    if len(sys.argv) != 5:
        print("Error: Incorrect number of parameters.")
        print("Usage: python topsis.py <inputFile> <weights> <impacts> <outputFile>")
        sys.exit(1)

    topsis(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])

def main():
    import sys

    if len(sys.argv) != 5:
        print("Usage: topsis <InputFile> <Weights> <Impacts> <OutputFile>")
        sys.exit(1)

    input_file = sys.argv[1]
    weights = sys.argv[2]
    impacts = sys.argv[3]
    output_file = sys.argv[4]

    topsis(input_file, weights, impacts, output_file)
