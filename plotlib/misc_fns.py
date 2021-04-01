from collections import OrderedDict
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from matplotlib.font_manager import FontProperties
from matplotlib.backends.backend_pdf import PdfPages
import pandas as pd
import numpy as np
import os

def initializePDFPage(filepath):
    """
    Return none if file already present.
    else return pdf_object.
    """

    if os.path.isfile(filepath) == True:
        msg = "ERROR: File already present at %s. PDF not created." % filepath
        print(msg)
        return None
    else:
        pdf_pages = PdfPages(filepath)
        return pdf_pages

def initialize_ax(ax=None, figsize=(8,6)):
    if ax == None:
        fig = plt.figure()
        fig.set_size_inches(*figsize)
        ax = fig.add_subplot(111)

    return ax

def reset_offset_for_xaxis(ax):
    
    ax.xaxis.get_major_formatter().set_useOffset(False)
    
    return ax

def set_visible_xticklabels(ax, visible=True):
    plt.setp(ax.get_xticklabels(),visible=visible)

def set_visible_yticklabels(ax, visible=True):
    plt.setp(ax.get_yticklabels(),visible=visible)

def set_visible_xlabels(ax, visible=True):
    ax.xaxis.label.set_visible(visible)

def set_visible_ylabels(ax, visible=True):
    ax.yaxis.label.set_visible(visible)

def set_axes_major_tick_size(ax, tickfontsize=10):
    # this does not work when changing size for twinned axis
    #if False:
    #    for tick in ax.xaxis.get_major_ticks():
    #        print "x before", tick.label1.get_fontsize()
    #        tick.label1.set_fontsize(tickfontsize)
    #        print "x after", tick.label1.get_fontsize()
    #    for tick in ax.yaxis.get_major_ticks():
    #        print "y before", tick.label1.get_fontsize()
    #        tick.label1.set_fontsize(tickfontsize)
    #        print "y after", tick.label1.get_fontsize()

    for tick in ax.get_xticklabels():
        tick.set_fontsize(tickfontsize)
    for tick in ax.get_yticklabels():
        tick.set_fontsize(tickfontsize)

def get_xticklabels_to_xticks(ax):
    #20170623 - initialised
    xticklabels_to_xticks = pd.Series(
        ax.get_xticks(),
        index=[t.get_text() for t in ax.get_xticklabels()]
        ).to_dict()
    return xticklabels_to_xticks 

def get_yticklabels_to_yticks(ax):
    #20170623 - initialised
    yticklabels_to_yticks = pd.Series(
        ax.get_yticks(),
        index=[t.get_text() for t in ax.get_yticklabels()]
        ).to_dict()
    return yticklabels_to_yticks 

def colorList_tableau(name):
    '''
    name = tableau10 | tableau10_light | tableau10_medium | tableau20
    http://tableaufriction.blogspot.ro/2012/11/finally-you-can-use-tableau-data-colors.html
    '''
    tableau10 = [
        (31,119,180), (255,127,14), (44,160,44), (214,39,40), (148,103,189),
        (140,86,75), (227,119,194), (127,127,127), (188,189,34), (23,190,207)]
    tableau10_light = [
        (174,199,232), (255,187,120), (152,223,138), (255,152,150), (197,176,213),
        (196,156,148), (247,182,210), (199,199,199), (219,219,141), (158,218,299)]
    tableau10_medium = [
        (114,158,206), (255,158,74), (103,191,92), (237,102,93), (173,139,201),
        (168,120,110), (237,151,202), (162,162,162), (205,204,93), (109,204,218)]
    tableau20 = [
        (31,119,180), (174,199,232), (255,127,14), (255,187,120), (44,160,44),
        (152,223,138), (214,39,40), (255,152,150), (148,103,189), (197,176,213),
        (140,86,75), (196,156,148), (227,119,194), (247,182,210), (127,127,127),
        (199,199,199), (188,189,34), (219,219,141), (23,190,207), (158,218,229)
        ]
    fn = lambda t: tuple([v/255.0 for v in t])
    name2colors = {
        'tableau10'         : [fn(i) for i in tableau10],
        'tableau10_light'   : [fn(i) for i in tableau10_light],
        'tableau10_medium'  : [fn(i) for i in tableau10_medium],
        'tableau20'         : [fn(i) for i in tableau20]
        }
    try:
        colors = name2colors[name]
    except:
        colors = name2colors['tableau20']
    return colors

def rebound(vallist, extensionratio = 0.05, ends="both"):
    """
    ends = both | left or bottom | right or top
    """
    minval = min(vallist)
    maxval = max(vallist)
    diff = maxval - minval
    extension = diff * float(extensionratio)
    extendedminval, extendedmaxval = minval, maxval
    ends = ends.lower()
    if ends == "left" or ends == "bottom":
        extendedminval = minval - extension
    elif ends == "right" or ends == "top":
        extendedmaxval = maxval + extension
    else:
        extendedminval = minval - extension
        extendedmaxval = maxval + extension

    return extendedminval, extendedmaxval

def get_new_limits_with_equal_segment_ratios(
    ref_min, ref_max, curr_min, curr_max):

    '''
    this function will use a reference ax and calculate
    the ratio of positive or negative segments,
    and calculate a new value for the new limits.

    #CHANGES
    #20160406 - initialized
    '''

    def calc_ratios(minv, maxv):
        assert (minv<=0) and (maxv>=0), 'no origin pt found.'
        span = float(maxv-minv)
        neg_ratio=abs(minv)/span
        pos_ratio=abs(maxv)/span
        #print "(%.3g, 0, %.3g) -> (%.3g | %.3g)" % (minv, maxv, neg_ratio, pos_ratio)
        return span, neg_ratio, pos_ratio

    #print 'ref:'
    ref_span, ref_neg_ratio, ref_pos_ratio = calc_ratios(ref_min, ref_max)

    #
    #print 'query:'
    curr_span, neg_ratio, pos_ratio = calc_ratios(curr_min, curr_max)

    if neg_ratio <= ref_neg_ratio:

        #
        new_max = curr_max
        new_min = curr_max-(curr_max/ref_pos_ratio)

    else:

        new_min = curr_min
        new_max = abs(new_min)/ref_neg_ratio + new_min

    #print 'after adjustment'
    new_span, new_neg_ratio, new_pos_ratio = calc_ratios(new_min, new_max)

    # check ratio
    assert customlib.compare2Vals(new_neg_ratio, ref_neg_ratio), '%s, %s' % (new_neg_ratio, ref_neg_ratio)
    assert customlib.compare2Vals(new_pos_ratio, ref_pos_ratio), '%s, %s' % (new_pos_ratio, ref_pos_ratio)
    assert curr_min >= new_min
    assert curr_max <= new_max

    return new_min, new_max

def remove_legend(ax):
    ax.legend_.remove()
    return ax
    
def set_visible_legend(ax, visible=True):
    ax.legend().set_visible(visible)