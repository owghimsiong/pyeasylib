'''
Helper classes for initialising figure object and fontset for plotting.

@author: oghimsio
'''


from misc_fns import *
from collections import OrderedDict
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from matplotlib.font_manager import FontProperties
import pandas as pd
import numpy as np
import pyeasylib
import pyeasylib.pdlib as pdlib

class FontSet:
    '''
    FigureFonts(self,
                labelsize = 18,
                ticksize = 12,
                titlesize = 18,
                legendsize = 18,
                textsize = 14,
                mainfamily="Arial",
                textboxfamily="monospace")
    '''
    def __init__(self,
                 labelsize = 18,
                 ticksize = 12,
                 titlesize = 18,
                 legendsize = 18,
                 textsize = 14,
                 mainfamily="Arial",
                 textboxfamily="monospace"):

        self.label = FontProperties(family=mainfamily, size=labelsize)
        self.tick = FontProperties(family=mainfamily, size=ticksize)
        self.title = FontProperties(family=mainfamily, size=titlesize)
        self.legend = FontProperties(family=mainfamily, size=legendsize)
        self.text = FontProperties(family=textboxfamily, size=textsize)

    def __repr__(self):

        attrs = [i for i in dir(self) if (not(i.startswith("__")) and not(i.endswith("__")))]
        df = pd.DataFrame(index=attrs, columns=['family', 'name', 'size'])
        df.index.name = "self.attr"
        for attr in attrs:
            attr_val = getattr(self, attr)
            name = attr_val.get_name()
            family = ", ".join(attr_val.get_family())
            size = attr_val.get_size()
            df.at[attr, 'family'] = family
            df.at[attr, 'name'] = name
            df.at[attr, 'size'] = size

        df_string = df.to_string()
        max_cols = max([len(l) for l in df_string.split("\n")])
        spacerline="-"*max_cols
        s = (
            "%s\n"
            "#Class: FontSet\n"
            "%s\n"
            "%s\n"
            "%s"
            ) % (spacerline, spacerline, df.to_string(), spacerline)

        return s

class InitializeFigureSubplots(object):
    '''
    InitializeFigureSubplots(
        self, axnames = [],
        num_ax_rows = True,
        num_ax_cols = True,
        axesnamedf=None,
        set_ax_title = True,
        width_ratios=None,
        height_ratios=None,
        figsize=[6,4],
        fontset = FontSet()
        )

    If axnames or axesnamedf contains np.nan, or None, they will
    not be plotted. If "None", they will be plotted.
    '''

    def __init__(self,
                 axnames = [],
                 num_ax_rows = True,
                 num_ax_cols = True,
                 axesnamedf=None, #
                 set_ax_title = True,
                 width_ratios=None,
                 height_ratios=None,
                 figsize=[6,4],

                 fontset = FontSet(),
                 ):

        import math
        NUMERICTYPES = pyeasylib.NUMERICTYPES

        # Added after porting to PY3.
        # because pd.isnull(range(5)) = False -> but will return a list in Py2
        # added a new line below as need it to return a list here.
        axnames = list(axnames)
        
        # either axesnamedf must be None or axnames must be []
        if (axesnamedf is not None) and (axnames not in [None, []]):
            msg = "ERROR: Both axnames and axesnamedf are provided - can only provide either one."
            raise Exception(msg)

        #added axesnamedf input
        if axesnamedf is None: # this means that axnames must be provided

            # added a check to see if empty list presetn
            num_axes = len(axnames)
            if num_axes == 0:
                raise Exception("Empty list found for axnames.")

            # num_mapping_dict - all non numeric to 1
            num_ax_rows = int(num_ax_rows) \
                          if type(num_ax_rows) in NUMERICTYPES \
                          else None
            num_ax_cols = int(num_ax_cols) \
                          if type(num_ax_cols) in NUMERICTYPES \
                          else None

            # convert boolean
            # both num not supplied
            if (num_ax_rows == None) and (num_ax_cols == None):
                #print 'here0'
                sqrt = num_axes ** 0.5
                int_sqrt = int(sqrt)
                if sqrt == int_sqrt:
                    num_ax_rows = int_sqrt
                    num_ax_cols = int_sqrt
                else:
                    num_ax_rows = int_sqrt
                    #num_ax_cols = int_sqrt + 1
                    num_ax_cols = int(math.ceil(float(num_axes) / int_sqrt))

            # only num cols is supplied
            elif (num_ax_rows == None) and (type(num_ax_cols) in NUMERICTYPES):
                num_ax_cols = int(num_ax_cols)
                num_ax_rows = int(math.ceil(float(num_axes) / num_ax_cols))
                #print 'here1'

            # only num rows is supplied
            elif (num_ax_cols == None) and (type(num_ax_rows) in NUMERICTYPES):
                num_ax_rows = int(num_ax_rows)
                num_ax_cols = int(math.ceil(float(num_axes) / num_ax_rows))
                #print 'here2'

            # both nums are supplied
            else:
                #print 'here3'
                # if here, means both input are
                pass

            # num calculated
            num_ax_calculated = num_ax_rows * num_ax_cols
            if num_ax_calculated < num_axes:
                error = ("Number of axes required = %s > "
                         "number of ax spaces available (%s row(s) x %s col(s) = %s).") % \
                        (num_axes, num_ax_rows, num_ax_cols, num_ax_calculated)
                raise Exception(error)

            #20170303 - added a axesnamedf
            axnames_tmp = list(axnames) + [None] * (num_ax_calculated-num_axes)
            axesnamedf = pd.DataFrame(
                np.reshape(axnames_tmp, (num_ax_rows, num_ax_cols)),
                index=pd.Index(list(range(num_ax_rows)), name="Row"),
                columns=pd.Index(list(range(num_ax_cols)), name="Column")
                )

        else: #axesnamedf is provided

            # get the dimension from the provided axesnamedf
            axesnamedf = axesnamedf.copy()
            axesnamedf.index = list(range(axesnamedf.shape[0]))
            axesnamedf.columns = list(range(axesnamedf.shape[1]))
            axnames = axesnamedf.values.flatten().tolist()
            num_ax_rows, num_ax_cols = axesnamedf.shape
            num_axes = num_ax_rows * num_ax_cols

        #check if non-null axnames are unique
        #modified to only check dupls for non None/np.Nan axnames
        #so that None can be used to fill in empty subplot axes
        non_null_axnames = np.array(axnames)[~pd.isnull(axnames)]
        #print "here"
        dup = pdlib.series_to_duplicates(non_null_axnames)
        if dup[0]:
            error = "Duplicate axnames found: %s" % dup[-1]
            raise Exception(error)

        # check the gridspec width ratios and height ratios
        if width_ratios != None:
            if len(width_ratios) == num_ax_cols:
                pass
            else:
                msg = "WARNING: num_ax_cols=%s but width_ratios=%s." % (num_ax_cols, width_ratios)
                width_ratios = [1] * num_ax_cols
                msg2 = "Setting to %s." % width_ratios
                print(msg + " " + msg2)
        else:
            width_ratios = [1] * num_ax_cols

        if height_ratios != None:
            if len(height_ratios) == num_ax_rows:
                pass
            else:
                msg = "WARNING: num_ax_rows=%s but height_ratios=%s." % (num_ax_rows, height_ratios)
                height_ratios = [1] * num_ax_rows
                msg2 = "Setting to %s." % height_ratios
                print(msg + " " + msg2)
        else:
            height_ratios = [1] * num_ax_rows

        # initialize location of ax df
        axesdf = pd.DataFrame(
            index=pd.Index(list(range(num_ax_rows)), name="Row"),
            columns=pd.Index(list(range(num_ax_cols)), name="Column"))
        axname2ax = OrderedDict()

        # fill the axes df
        fig = plt.figure(1)
        fig.clear()
        fig.set_size_inches(*figsize)
        gs = gridspec.GridSpec(num_ax_rows, num_ax_cols,
                               width_ratios=width_ratios,
                               height_ratios=height_ratios)

        #rewritten so that we can reference from axesnamedf
        for i in axesnamedf.index:
            for c in axesnamedf.columns:
                subplotindex = (i*num_ax_cols) + (c)
                axname = axesnamedf.at[i, c]
                if not(pd.isnull(axname)): 
                    #print i, c, subplotindex, axname
                    ax = fig.add_subplot(gs[subplotindex])
                    if set_ax_title:
                        ax.set_title(axname)
                    axesdf.iat[i, c] = ax
                    axname2ax[axname] = ax

        #
        self.axnames = list(axnames)
        self.num_axes = num_axes
        self.num_ax_rows = num_ax_rows
        self.num_ax_cols = num_ax_cols
        self.axesdf = axesdf
        self.axesnamedf = axesnamedf
        self.axname2ax = axname2ax
        self.figure = fig
        self.fontset = fontset

    #
    def get_ax(self, axname):
        return self.axname2ax[axname]

    def get_uninitialized_ax(self):
        uninitialized_axes =  pdlib.df_find_value(self.axesnamedf, np.nan)
        return uninitialized_axes

    def get_all_axes(self):
        return list(self.axname2ax.values())

    def get_left_col_axes(self):
        axes = list(self.axesdf.iloc[:, 0].values)
        return axes

    def get_nonleft_col_axes(self):
        axes = list(self.axesdf.iloc[:, 1:].values.flatten())
        return axes

    def get_right_col_axes(self):
        axes = list(self.axesdf.iloc[:, -1].values)
        return axes

    def get_nonright_col_axes(self):
        axes = list(self.axesdf.iloc[:, :-1].values.flatten())
        return axes

    def get_bottom_row_axes(self):
        axes = list(self.axesdf.iloc[-1, :].values)
        return axes

    def set_left_col_axes_label(self, label):
        for ax in self.get_left_col_axes():
            ax.set_ylabel(label, fontproperties=self.fontset.label)

    def set_bottom_row_axes_label(self, label):
        for ax in self.get_bottom_row_axes():
            ax.set_xlabel(label, fontproperties=self.fontset.label)

    def savefig(self, fp):
        self.figure.savefig(fp)

    def savefig_to_pdfobj(self, pdfobj):
        pdfobj.savefig(self.figure)

    def tight_layout(self, **kwargs):
        self.figure.tight_layout(**kwargs)