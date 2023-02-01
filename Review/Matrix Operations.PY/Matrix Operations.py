"""
#==========================================================================
Python-Script for Vissim 9+
Copyright (C) PTV AG. All rights reserved.
Jochen Lohmiller 2017
-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
Example for matrix operations
#==========================================================================

This script demonstrates matrix operations such as:
- creating a matrix
- reading and setting matrix values
- change the time period of a matrix
- apply the matrix to the dynamic assignment demand
- importing an *.fma file as a matrix
- removing a matrix
"""

# COM-Server
import os
import numpy

def StartVissim():
    # COM-Server
    import win32com.client as com
    ## Connecting the COM Server => Open a new Vissim Window:
    Vissim = com.Dispatch("Vissim.Vissim")
    return Vissim


def toList(NestedTuple):
    """
    function to convert a nested tuple to a nested list
    """
    return list(map(toList, NestedTuple)) if isinstance(NestedTuple, (list, tuple)) else NestedTuple


def getMatrix(matrix):
    """
    reading an entire matrix
    returns a numpy array
    """
    numberOfColumns = matrix.ColCount
    # numberOfRows= newMatrix.RowCount
    # note: matrices are always symmetric with the dimension of number of zones
    matrixValues = list([])
    for column in range(numberOfColumns):
        matrixValues.append(list())
        for row in range(numberOfColumns):
            matrixValues[column].append(list())
            matrixValues[column][row] = matrix.GetValue(column + 1, row + 1) # plus 1 because Vissim starts with 1 (Pyhton with 0)
    matrixValuesNP = numpy.array(matrixValues)
    return matrixValuesNP

def setMatrix(matrix, matrixNewValues):
    """
    reading an entire matrix
    returns a numpy array
    setting the entire matrix:
    """
    numberOfColumns = matrix.ColCount
    # numberOfRows= newMatrix.RowCount
    # note: matrices are always symmetric with the dimension of number of zones
    for column in range(numberOfColumns):
        for row in range(numberOfColumns):
            matrix.SetValue(column + 1, row + 1, matrixNewValues[column, row]) # plus 1 because Vissim starts with 1 (Python with 0)


def matrixOperations(Vissim):

    Vissim.SuspendUpdateGUI()   # stop updating of the complete Vissim workspace (network editor, list, chart and signal time table windows)

    # assign a name to the existing matrix
    Vissim.Net.Matrices.ItemByKey(1).SetAttValue('Name', 'Original matrix')

    # add a new matrix
    newMatrix = Vissim.Net.Matrices.AddMatrix(0)

    # get the number of columns and rows
    numberOfColumns = newMatrix.ColCount
    # numberOfRows= newMatrix.RowCount
    # note: matrices are always symmetric with the dimension of number of zones

    # read the demand of the matrix
    demand12 = newMatrix.GetValue(1, 2)

    # read an entire matrix
    matrixValuesNP = getMatrix(newMatrix)

    # set the value of the matrix
    newMatrix.SetValue(1, 2, 550)

    # set the entire matrix
    matrixValuesNP[0, 1] = 500
    matrixValuesNP[1, 0] = 370
    setMatrix(newMatrix, matrixValuesNP)

    # set the time
    newMatrix.SetAttValue('FromTime', '00:00:00')
    newMatrix.SetAttValue('ToTime', '00:45:00')
    # note: matrix values are not vehicles per hour!
    # matrix values are vehicles per time period, example: a matrix value of 500 with a time from 00:00:00 - 00:45:00 will generate 500 vehicles in 45 minutes.

    # set a name
    newMatrix.SetAttValue('Name', 'My new updated matrix')

    # import a matrix from *.fma file
    fmaMatrix = Vissim.Net.Matrices.AddMatrix(0) # first add another matrix
    currentPath = os.getcwd()
    fmaFileName = 'Demand.fma'
    filePath = os.path.join(currentPath, fmaFileName)
    ReadMatrixfromFile(fmaMatrix, filePath)

    # assign the matrix in the dynamic assignment settings
    usedMatrices = toList(Vissim.Net.DynamicAssignment.DynAssignDemands.GetMultiAttValues('Matrix'))
    usedMatrices[0][1] = fmaMatrix # does not necessarily be the object, string or integer of ID works also: usedMatrices[0][1] = '2' or usedMatrices[0][1] = 2
    Vissim.Net.DynamicAssignment.DynAssignDemands.SetMultiAttValues('Matrix', usedMatrices)

    # alternative method for assigning a matrix: use remove/add
    # remove:
    Vissim.Net.DynamicAssignment.DynAssignDemands.RemoveDynAssignDemand(Vissim.Net.DynamicAssignment.DynAssignDemands.ItemByKey(1))
    # add:
    DynAssignDemand = Vissim.Net.DynamicAssignment.DynAssignDemands.AddDynAssignDemand(0)
    DynAssignDemand.SetAttValue('Matrix', fmaMatrix)

    # run simulation (4 runs as configured in simulation parameters); mainly in order to get found paths from the dynamic assignment
    Vissim.Simulation.RunContinuous()

    # configure matrix correction
    da = Vissim.Net.DynamicAssignment
    da.SetAttValue("MATRIX", 2) # set the ID of the matrix that shall be used, in our case we use the 2nr matrix that we made above
    cnt = da.CntDataAttr # this represents the list of attributes of links for "Counts for Links" from the matrix correction dialog    
    # Now lets set the count data for the matrix correction. We want to aggregate counts on links for vehicle classes 10, 20, 30.
    # Unfortunately, we do have a "bug" in the COM interface here: instead of just providing a list of link attribute names, the interface
    # requires us to provide some additional parameters. So instead of our desired <attributeID>, we have to provide the following tuple 
    # of values <attributeID>, <numberOfDecimalPlaces>, <formatSpecifier>, <showUnitSymbol> for each attribute we want to aggregate.
    # The values for the additional arguments do not affect the outcome of the matrix correction. We just use arbitrary values 2, 1, 1 for those.
    cnt.ReplaceAll(["CNTDATA(10)", 2, 1, 1, "CNTDATA(20)", 2, 1, 1, "CNTDATA(30)", 2, 1, 1])
    vols = da.PathVolAttr # this represents the list of attributes of links for "Volumes on Paths" from the matrix correction dialog
    # As volume attributes for the matrix correction, we want to aggregate the path volumes of all vehicle classes 
    # over the first six time intervals.
    # Again, we need to provide the additional arguments when setting path volumes. 
    vols.ReplaceAll(["VOLNEW(1,ALL)", 2, 1, 1, "VOLNEW(2,ALL)", 2, 1, 1, "VOLNEW(3,ALL)", 2, 1, 1, "VOLNEW(4,ALL)", 2, 1, 1, "VOLNEW(5,ALL)", 2, 1, 1, "VOLNEW(6,ALL)", 2, 1, 1])

    # read the demand of the matrix before correction
    demand12Before = newMatrix.GetValue(1, 2)

    # run matrix correction
    da.RunMatrixCorrection()

    # read the demand of the matrix after correction
    demand12After = newMatrix.GetValue(1, 2)

    # add another matrix
    nextMatrix = Vissim.Net.Matrices.AddMatrix(0)

    # initialize a matrix: Set all matrix values to zero
    nextMatrix.Init()

    # remove a matrix
    Vissim.Net.Matrices.RemoveMatrix(nextMatrix)

    Vissim.ResumeUpdateGUI(True)    # allow updating of the complete Vissim workspace (network editor, list, chart and signal time table windows)

def ReadMatrixfromFile(matrix, filePath):
    """
    Imports a matrix from a *.fma files. fma files were used for matrices in PTV Vissim before version 9.
    """
    matrix.ReadFromFile(filePath)
    matrix.SetAttValue('Name', 'Import from \'' + filePath.split('\\')[-1] + '\'') # filePath.split('\\')[-1] is the filename without the full path

def main():
    """
    Main control
    """
    Vissim = StartVissim()
    # Load a Vissim Network:
    currentPath = os.getcwd()
    network = '3 Paths.inpx'
    Filename = os.path.join(currentPath, network)
    Vissim.LoadNet(Filename)

    # Save under different file name:
    networkSaved = '3 Paths saved.inpx'
    Filename = os.path.join(currentPath, networkSaved)
    Vissim.SaveNetAs(Filename)

    matrixOperations(Vissim)

def mainWithoutStart():
	# function for internal script
    currentPath = os.getcwd()

    # Save under different file name:
    networkSaved = '3 Paths saved.inpx'
    Filename = os.path.join(currentPath, networkSaved)
    Vissim.SaveNetAs(Filename)

    Vissim.SetAttValue('ShowMessages', True)
    matrixOperations(Vissim)

if __name__ == '__main__':
    main()
