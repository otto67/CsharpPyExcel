from genericpath import exists
import sys, os
# Reuse simulator
import Simulator
import numpy as np
import pathlib


def readFromFile(filename):
    """[summary]
    Reads input data for simulator from file

    Args:
        filename ([type]): File to be read

    Returns:
        List: List of values being input to simulator  
    """
    return_list = []
    try: 
        file = open(filename, 'r')
        lines = file.readlines()
        validinputs = ('right hand side','number of nodes', 'domain','method of solution', 'rhs_coeff', 'bclow', 'bcupp', 'bcleft', 'bcright')
        for line in lines:
            if line.strip()[0] == '#':  # This is a comment
                continue
            
            ans = line.find(':')
            if ans < 0:
                continue
            
            entry = line[:ans].strip()
            if entry.lower() in validinputs:
                return_list.append(line.strip())

    except Exception as e:
        print("Exception occurred!", e)
    finally:
        return return_list

def solve():
    """ Runs simulator 
    """   
    fname = os.path.join(os.path.dirname(os.path.abspath(__file__)),'input.txt')

    x_min = y_min = x_max = y_max = 0

    try:
        if os.path.isfile(fname):
            f = open(fname, "r")
            mylist = readFromFile(fname)
    
        # Extract domain for use when writing results to file    
        for input in mylist:
            if input.split(':')[0].lower() == "domain":
                (x_min, x_max, y_min, y_max) = Simulator.parse_domain(input.split(':')[1])
    
        if (x_max == 0):
            x_max = 1
            x_max = 0
        if (y_max == 0):
            y_max = 1
            y_min = 0

        # Test function. To be removed 
        # writeToFile()
        
        sol = Simulator.run(mylist)
        writeNodeValsToFile(sol, x_max, x_min, y_max, y_min)
    except Exception as e:
        print("Error in reading file,",e)
    

def writeNodeValsToFile(solution, xmax, xmin, ymax, ymin):
    
    """ Write node values to file 

    Args:
        solution (2d numpy array): Nodal values of solution
        xmax (float): Right boundary
        xmin (float): Left boundary
        ymax (float): Upper boundary
        ymin (float): Lower boundary
    """    
    filename = os.path.join(os.path.dirname(os.path.abspath(__file__)),'nodevals.txt')

    dx = (float) ((xmax - xmin)/(solution.shape[0]-1))
    dy = (float) ((ymax - ymin)/(solution.shape[1]-1))

    try: 
        f = open(filename, "w")

        tmp = str(solution.shape[0]) # upper left corner. 
        for i in range(solution.shape[0]):
            x = xmin + i*dx
            tmp += f",{x}"
        f.write(tmp + '\n')
        
        
        for i in range(solution.shape[0]):
            yval = ymin + i*dy
            f.write(f"{yval}")  
            for j in range(solution.shape[1]):                 
                f.write(f",{solution[i,j]}")
            f.write("\n") 
        
        f.close()    

    except Exception as e:
        print("Exception caught", e)


def myFunc(x, y):
    return x*y


def writeToFile():
    filename = 'nodevals.txt'

    f = open(filename, "w")

    dx = dy = 0.1
    x = 0.0

    # First line is x-values
    tmp = "21" # upper left corner - unused
    for i in range(21):
        x = -1.0 + i*dx
        tmp += f",{x}"
    f.write(tmp + '\n')
    # First column is y-values
    x = 0.0
    y = 0.0
    for i in range(21):
        x = -1.0 + i*dx
        f.write(f"{x}") # NOTE: this assumes dx = dy 
        for j in range(21):             
            y = -1.0 + j*dy       
            f.write(f",{myFunc(x,y)}")
        f.write("\n")
    
    f.close()

if (__name__ == "__main__"):
    solve() 