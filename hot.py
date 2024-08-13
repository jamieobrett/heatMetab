############################################################################
### Functions ###

#https://stackoverflow.com/questions/16834861/create-own-colormap-using-matplotlib-and-plot-color-scale
#Info about the returned object: https://matplotlib.org/stable/api/_as_gen/matplotlib.colors.LinearSegmentedColormap.html#matplotlib.colors.LinearSegmentedColormap, https://matplotlib.org/stable/tutorials/colors/colormap-manipulation.html
#To use the final cmap returned, use it as cmap(normalized-float) with the input normalized to be 0 to 1 and with a decimal point in it
import matplotlib.pyplot
import matplotlib.colors
import numpy as np
def colorMapMaker(valueToColorDict,colorbarImagePrefix='colorbar'):
    """Requires matplotlib.pyplot and matplotlib.colors. Note that the numbers entered must be floats. Returns a matplotlib.colors.Colormap object that takes an already-normalized (range 0 to 1) float (not int or string) and returns an (R,G,B,alpha) tuple"""
    num_smallest = min(valueToColorDict.keys())
    num_highest = max(valueToColorDict.keys())
    norm = matplotlib.pyplot.Normalize(num_smallest,num_highest)
    tuples = list(zip(map(norm,valueToColorDict.keys()), valueToColorDict.values()))
    cmap = matplotlib.colors.LinearSegmentedColormap.from_list("", tuples)
    #make the legend
    fig, ax = matplotlib.pyplot.subplots()
    cax = ax.imshow([[0,1]], cmap=cmap)
    NTICKS = 4
    floatrange = list(np.linspace(0,1,NTICKS+1,endpoint=1)) #for the colorbar (whose actual range is 0 to 1 of normalized values)
    mylabels = list(np.linspace(num_smallest,num_highest,NTICKS+1,endpoint=1))
    cbar = fig.colorbar(cax, ticks=floatrange)
    cbar.ax.set_yticklabels(mylabels)
    ax.set_visible(False)
    matplotlib.pyplot.savefig(colorbarImagePrefix + '.png', bbox_inches='tight')
    matplotlib.pyplot.savefig(colorbarImagePrefix + '.pdf', bbox_inches='tight')
    matplotlib.pyplot.savefig(colorbarImagePrefix + '.svg', bbox_inches='tight')
    return cmap

def normalizeTheseCoordinates(coordList,boxwidth,boxheight,offsetx,offsety):
    """Takes a graph_force output of a list of [(xy-tuple coordinates)] and normalizes it to span a certain box size"""
    xmin = min([t[0] for t in coordList])
    xmax = max([t[0] for t in coordList])
    ymin = min([t[1] for t in coordList])
    ymax = max([t[1] for t in coordList])
    boxheight = boxheight-offsety
    boxwidth = boxwidth-offsetx
    normalizedCoordList = [((t[0]-xmin)/(xmax-xmin)*(boxwidth) + offsetx,(t[1]-ymin)/(ymax-ymin)*(boxheight) + offsety) for t in coordList]
    return normalizedCoordList

############################################################################
### Classes ###

class Metabolite:
    """A class for metabolites"""
    def __init__(self, name, fc, q):
        self.name = name.upper()
        self.nodeIndices = [] #can appear multiple times, or not at all
        self.fc = fc
        self.q = q
        self.outlineCol = None
        self.fillCol = None
        self.displayName = name #default if none provided
        self.edgeCount = 0
    #Method to populate the outlineCol and fillCol given a colormap made by colorMapMaker, the dataFloor and dataCeil to normalize this metabolite's value to, and the FDR threshold desired
    def gimmeColors(self, colormap, dataFloor, dataCeil, fdrThresh, undetectedColor):
        if self.fc is None:
            self.outlineCol = matplotlib.colors.to_rgba(undetectedColor)
            self.fillCol =  (1,1,1,0) #transparent
        else:
            normalizedFCFloat = float (self.fc - dataFloor) / (dataCeil - dataFloor)
            myCol = colormap(normalizedFCFloat)
            self.outlineCol = myCol
            if self.q > fdrThresh:
                self.fillCol = (1,1,1,0) #transparent
            else:
                self.fillCol = myCol

class Edge:
    """A class for metabolite edges"""
    def __init__(self, node1, node2):
        self.edgeTuple = (node1, node2) #node indices, not metabolite names

class Node:
    """"A class for metabolite nodes"""
    def __init__(self,name,nodeIndex,center_x,center_y):
        self.name = name
        self.nodeIndex = nodeIndex
        self.center_x = center_x
        self.center_y = center_y