'''
Minimal script to plot simulation results from PSCAD and PowerFactory.
'''
from __future__ import annotations
from os import listdir, makedirs
from os.path import abspath, join, split, exists
import re
import pandas as pd
from plotly.subplots import make_subplots  # type: ignore
import plotly.graph_objects as go  # type: ignore
from typing import List, Dict, Union, Tuple, Set
from sampling_functions import downSample
from threading import Thread, Lock
import time
import sys
from math import ceil
from read_configs import ReadConfig, readFigureSetup, readCursorSetup
from Figure import Figure
from Result import ResultType, Result
from Cursor import Cursor
from read_and_write_functions import loadEMT
from process_results import getColNames, getUniqueEmtSignals
from process_psout import getSignals
from cursor_functions import setupCursorDataFrame, addCursorMetrics
from ideal_functions import genIdealResults

try:
    LOG_FILE = open('plotter.log', 'w')
except:
    print('Failed to open log file. Logging to file disabled.')
    LOG_FILE = None  # type: ignore

gLock = Lock()


def print(*args):  # type: ignore
    '''
    Overwrites the print function to also write to a log file.
    '''
    gLock.acquire()
    outputString = ''.join(map(str, args)) + '\n'  # type: ignore
    sys.stdout.write(outputString)
    if LOG_FILE:
        try:
            LOG_FILE.write(outputString)
            LOG_FILE.flush()
        except:
            pass
    gLock.release()


def idFile(filePath: str) -> Tuple[
    Union[ResultType, None], Union[int, None], Union[str, None], Union[str, None], Union[str, None]]:
    '''
    Identifies the type (EMT or RMS), root and case id of a given file. If the file is not recognized, a none tuple is returned.
    '''
    path, fileName = split(filePath)
    match = re.match(r'^(\w+?)_([0-9]+).(inf|csv|psout|zip|gz|bz2|xz)$', fileName.lower())
    if match:
        rank = int(match.group(2))
        projectName = match.group(1)
        bulkName = join(path, match.group(1))
        fullpath = filePath
        if match.group(3) == 'psout':
            fileType = ResultType.EMT_PSOUT
            return (fileType, rank, projectName, bulkName, fullpath)
        elif match.group(3) == 'zip' or match.group(3) == 'gz' or match.group(3) == 'bz2' or match.group(3) == 'xz':
            fileType = ResultType.EMT_ZIP
            return (fileType, rank, projectName, bulkName, fullpath)
        else:
            with open(filePath, 'r') as file:
                firstLine = file.readline()
                if match.group(3) == 'inf' and firstLine.startswith('PGB(1)'):
                    fileType = ResultType.EMT_INF
                    return (fileType, rank, projectName, bulkName, fullpath)
                elif match.group(3) == 'csv' and firstLine.startswith('time;'):
                    fileType = ResultType.EMT_CSV
                    return (fileType, rank, projectName, bulkName, fullpath)
                elif match.group(3) == 'csv':
                    secondLine = file.readline()
                    if secondLine.startswith(r'"b:tnow in s"'):
                        fileType = ResultType.RMS
                        return (fileType, rank, projectName, bulkName, fullpath)
                
    return (None, None, None, None, None)


def mapResultFiles(config: ReadConfig) -> Dict[int, List[Result]]:
    '''
    Goes through all files in the given directories and maps them to a dictionary of cases.
    '''
    files: List[Tuple[str, str]] = list()
    for dir_ in config.simDataDirs:
        for file_ in listdir(dir_[1]):
            files.append((dir_[0], join(dir_[1], file_)))
    
    results: Dict[int, List[Result]] = dict()

    for file in files:
        group = file[0]
        fullpath = file[1]
        typ, rank, projectName, bulkName, fullpath = idFile(fullpath)

        if typ is None:
            continue
        assert rank is not None
        assert projectName is not None
        assert bulkName is not None
        assert fullpath is not None

        newResult = Result(typ, rank, projectName, bulkName, fullpath, group)

        if rank in results.keys():
            results[rank].append(newResult)
        else:
            results[rank] = [newResult]

    return results


def colorMap(results: Dict[int, List[Result]]) -> Dict[str, List[str]]:
    '''
    Select colors for the given projects. Return a dictionary with the project name as key and a list of colors as value.
    '''
    colors = ['#e6194B', '#3cb44b', '#ffe119', '#4363d8', '#f58231', '#911eb4', '#42d4f4', '#f032e6', '#bfef45',
              '#fabed4', '#469990', '#dcbeff', '#9A6324', '#fffac8', '#800000', '#aaffc3', '#808000', '#ffd8b1',
              '#000075', '#a9a9a9', '#000000']

    projects: Set[str] = set()

    for rank in results.keys():
        for result in results[rank]:
            projects.add(result.shorthand)

    cMap: Dict[str, List[str]] = dict()

    if len(list(projects)) > 2:
        i = 0
        for p in list(projects):
            cMap[p] = [colors[i % len(colors)]] * 3
            i += 1
        return cMap
    else:
        i = 0
        for p in list(projects):
            cMap[p] = colors[i:i + 3]
            i += 3
    cMap['ideal'] = ['#02525e', '#028b76', '#00847c']
    
    return cMap


def addResults(plots: List[go.Figure],
               result, # result object
               resultData: pd.DataFrame,
               figures: List[Figure],
               colors: Dict[str, List[str]],
               nColumns: int,
               pfFlatTIme: float,
               pscadInitTime: float,
               settingDict, # project settings
               caseDf # case MTB setting
               ) -> None:
    '''
    Add results for a case/rank to to a set of plots or a single subplot.
    '''

    SUBPLOT = (len(plots) == 1) # Check if output should be a subplot
    
    ideal = genIdealResults(result, resultData, settingDict,  caseDf, pscadInitTime)
    
    rowPos = 1
    colPos = 1
    
    fi = -1
    for figure in figures:
        fi += 1

        if not SUBPLOT: # Make use of individual plots
            plotlyFigure = plots[fi]
        else:           # Make use of subplots
            plotlyFigure = plots[0]
            rowPos = (fi // nColumns) + 1
            colPos = (fi % nColumns) + 1

        downsampling_method = figure.down_sampling_method
        timeColName = 'time' if result.typ in (ResultType.EMT_INF, ResultType.EMT_PSOUT, ResultType.EMT_CSV, ResultType.EMT_ZIP) else resultData.columns[0]
        timeoffset = pfFlatTIme if result.typ == ResultType.RMS else pscadInitTime

        # Add ideal result plots
        if figure.title in ideal['figs']:
            i = ideal['figs'].index(figure.title) # Get the index value for the figure.title in the ideal results
            x_value = ideal['data']['time']
            y_value = ideal['data'][ideal['signals'][i]]
            x_value, y_value = downSample(x_value, y_value, downsampling_method, figure.gradient_threshold)
            add_scatterplot_for_result(colPos, colors, 'dash', 'ideal:'+ideal['signals'][i], SUBPLOT, plotlyFigure, 'ideal', rowPos,
                                       0, x_value, y_value)
            
        traces = 0
        for sig in range(1, 4):
            signalKey = result.typ.name.lower().split('_')[0]
            rawSigName: str = getattr(figure, f'{signalKey}_signal_{sig}')
            sigColName, sigDispName= getColNames(rawSigName,result)

            if sigColName in resultData.columns:
                x_value = resultData[timeColName] - timeoffset  # type: ignore
                y_value = resultData[sigColName]  # type: ignore
                x_value, y_value = downSample(x_value, y_value, downsampling_method, figure.gradient_threshold)
                add_scatterplot_for_result(colPos, colors, 'solid', sigDispName, SUBPLOT, plotlyFigure, result.shorthand, rowPos,
                                           traces, x_value, y_value)

                # plot_cursor_functions.add_annotations(x_value, y_value, plotlyFigure)
                traces += 1
            elif sigColName != '' and result.typ != ResultType.EMT_CSV: # Temporary fix for ideal output result type files where not all signals are present
                print(f'Signal "{rawSigName}" not recognized in resultfile: {result.fullpath}')
                add_scatterplot_for_result(colPos, colors, 'solid', f'{sigDispName} (Unknown)', SUBPLOT, plotlyFigure, result.shorthand, rowPos,
                                           traces, None, None)
                traces += 1
        
        update_y_and_x_axis(colPos, figure, nColumns, plotlyFigure, rowPos)


def update_y_and_x_axis(colPos, figure, nColumns, plotlyFigure, rowPos):
    if nColumns in (1,2,3):
        yaxisTitle = f'[{figure.units}]'
    else:
        yaxisTitle = f'{figure.title}[{figure.units}]'
    if nColumns in (1,2,3):
        plotlyFigure.update_xaxes(  # type: ignore
            title_text='Time[s]'
        )
        plotlyFigure.update_yaxes(  # type: ignore
            title_text=yaxisTitle
        )
    else:
        plotlyFigure.update_xaxes(  # type: ignore
            title_text='Time[s]',
            row=rowPos, col=colPos
        )
        plotlyFigure.update_yaxes(  # type: ignore
            title_text=yaxisTitle,
            row=rowPos, col=colPos
        )


def add_scatterplot_for_result(colPos, colors, dash, displayName, SUBPLOT, plotlyFigure, resultName, rowPos, traces, x_value,
                               y_value):
    if not SUBPLOT:
        plotlyFigure.add_trace(  # type: ignore
            go.Scatter(
                x=x_value,
                y=y_value,
                line=dict(dash=dash),
                name=displayName,
                legendgroup=displayName,
                showlegend=True
            )
        )
    else:
        plotlyFigure.add_trace(  # type: ignore
            go.Scatter(
                x=x_value,
                y=y_value,
                #line_color=colors[resultName][traces],
                name=displayName,
                legendgroup=resultName,
                showlegend=True
            ),
            row=rowPos, col=colPos
        )


def genCursorPlotlyTables(ranksCursor, dfCursorsList):
    goCursorList = []
    for i, cursor in enumerate(ranksCursor):
        cursor_title = cursor.title
        rows = len(dfCursorsList[i])
        fig = go.Figure(data=[go.Table(header=dict(values=list(dfCursorsList[i].columns),
                                                   fill_color='#00847c',
                                                   font=dict(color='#ffffff'),
                                                   align='left'),
                                       cells=dict(values=[dfCursorsList[i][f'{column}'] for column in dfCursorsList[i].columns],
                                                  fill_color='#d8d8d8',
                                                  font=dict(color='#02525e'),
                                                  align='left'))
                              ])
        fig.update_layout(title_text=cursor_title)
        fig.update_layout(height = (rows+1)*50, margin=dict(t=50, l=50, r=50, b=0))
        goCursorList.append(fig)
        
    return goCursorList

    
def genCursorPdf(rank, rankName, ranksCursor, dfCursorsList, nColumns, figurePath):
    row_col_specs = []
    rows=ceil(len(ranksCursor)/nColumns)
    for i in range(rows):
        row_spec =[]
        for j in range(nColumns):
            row_spec.append({"type": "table"}) # default for 1 column
        row_col_specs.append(row_spec)        
    fig = make_subplots(
            rows=ceil(len(ranksCursor)/nColumns), cols=nColumns,
            vertical_spacing=0.03,
            specs=row_col_specs)
    total_number_of_rows = 0
    for i, cursor in enumerate(ranksCursor):
        total_number_of_rows += len(dfCursorsList[i])
        fig.add_trace(go.Table(header=dict(values=list(dfCursorsList[i].columns),
                                           fill_color='#00847c',
                                           font=dict(size=10, color='#ffffff'),
                                           align='left'),
                               cells=dict(values=[dfCursorsList[i][f'{column}'] for column in dfCursorsList[i].columns],
                                          fill_color='#d8d8d8',
                                          font=dict(size=10, color='#02525e'),
                                          align='left'),
                               ),
                               row=ceil((i+1)/nColumns),
                               col=(i % nColumns)+1
                     )
    fig.update_layout(
        showlegend=False,
        title_text=f"Cursor Metric Data for Rank {rank}: {rankName}",
        margin=dict(t=50, l=50, r=50, b=50)
        )
    cursor_path = figurePath + "_cursor"
    fig.write_image(f'{cursor_path}.pdf', height=50*total_number_of_rows, width=800*nColumns)    

    
def genCursorHTML(htmlCursorColumns, goCursorList):
    html = '<h2><div id="Cursors">Cursors:</div></h2><br>'
    html += '<div style="text-align: left; margin-top: 1px;">'
    for goCursor in goCursorList:
        cursor_title = goCursor['layout']['title']['text']
        cursor_ref = cursor_title.replace(' ', '_')
        html += f'<a href="#{cursor_ref}">{cursor_title}</a>&emsp;'
    html += '</div>'
    html += '<table style="width:100%">'
    html += '<tr>'
    for i in range(htmlCursorColumns):
        html += f'<th style"width:{round(100/htmlCursorColumns)}%"> &nbsp; </th>'
    html += '<tr>'
    for i, goCursor in enumerate(goCursorList):
        cursor_title = goCursor['layout']['title']['text']
        cursor_ref = cursor_title.replace(' ', '_')
        if ((i+1) % htmlCursorColumns) == 1:
            html += '<tr>'
        html += f'<td><div id="{cursor_ref}">' + goCursor.to_html(full_html=False,
                                                                         include_plotlyjs='cdn',
                                                                         include_mathjax='cdn',
                                                                         default_width='100%') + '</div></td>'  # type: ignore
        if ((i+1) % htmlCursorColumns) == 0:
            html += '</tr>'
    html += '</table>'
    
    return html


def drawPlot(rank: int,
             resultDict: Dict[int, List[Result]],
             figureDict: Dict[int, List[Figure]],
             casesDf, # Pandas DataFrame
             colorMap: Dict[str, List[str]],
             cursorDict: List[Cursor],
             settingDict: Dict[str, str],
             config: ReadConfig):
    '''
    Draws plots for html and static image export.    
    '''
    caseDf = casesDf[casesDf['Case']['Rank']==rank] # Get all case data for the current rank
    rankName = caseDf['Case']['Name'].squeeze()     # Get the rank Name for the current rank
    
    print(f'Drawing plot for Rank {rank}: {rankName}')

    resultList = resultDict.get(rank, [])
    rankList = list(resultDict.keys())
    rankList.sort()
    figureList = figureDict[rank]
    ranksCursor = [i for i in cursorDict if i.id == rank]

    if resultList == [] or figureList == []:
        return

    figurePath = join(config.resultsDir, str(rank))

    htmlPlots: List[go.Figure] = list()
    imagePlots: List[go.Figure] = list()

    setupPlotLayout(rankName, config, figureList, htmlPlots, imagePlots, rank)
    if len(ranksCursor) > 0:
        dfCursorsList = setupCursorDataFrame(ranksCursor)
    for result in resultList:
        print(f'Processing: {result.fullpath}')
        if result.typ == ResultType.RMS:
            resultData: pd.DataFrame = pd.read_csv(result.fullpath, sep=';', decimal=',', header=[0, 1])  # type: ignore
        elif result.typ == ResultType.EMT_INF:
            resultData: pd.DataFrame = loadEMT(result.fullpath)
        elif result.typ == ResultType.EMT_PSOUT:
            resultData: pd.DataFrame = getSignals(result.fullpath, getUniqueEmtSignals(figureList))
        elif result.typ == ResultType.EMT_CSV or result.typ == ResultType.EMT_ZIP:
            resultData: pd.DataFrame = pd.read_csv(result.fullpath, sep=';', decimal=',')  # type: ignore
        else:
            continue

        if config.genHTML:
            addResults(htmlPlots, result, resultData, figureList, colorMap,
                       config.htmlColumns, config.pfFlatTIme, config.pscadInitTime, settingDict, caseDf)
        if config.genImage:
            addResults(imagePlots, result, resultData, figureList, colorMap,
                       config.imageColumns, config.pfFlatTIme, config.pscadInitTime, settingDict, caseDf)
        if len(ranksCursor) > 0:
            addCursorMetrics(ranksCursor, dfCursorsList, result, resultData, config.pfFlatTIme, config.pscadInitTime, settingDict,  caseDf)
    
    goCursorList = genCursorPlotlyTables(ranksCursor, dfCursorsList)
    
    if config.genHTML:
        create_html(htmlPlots, goCursorList, figurePath, rankName if rankName is not None else "", rank, config, rankList)
        print(f'Exported plot for Rank {rank} to {figurePath}.html')

    if config.genImage:
        # Cursor plots are not currently supported for image export and commented out
        create_image_plots(config, figureList, figurePath, imagePlots)
        genCursorPdf(rank, rankName, ranksCursor, dfCursorsList, config.imageCursorColumns, figurePath)
    
        # create_cursor_plots(config.htmlCursorColumns, config, figurePath, imagePlotsCursors, ranksCursor)
        print(f'Exported plot for Rank {rank} to {figurePath}.{config.imageFormat}')

    print(f'Plot for Rank {rank} done.')


def create_image_plots(config, figureList, figurePath, imagePlots):
    if config.imageColumns == 1:
        # Combine all figures into a single plot, same as for nColumns > 1 but no grid needed
        combined_plot = make_subplots(rows=len(imagePlots), cols=1,
                                      subplot_titles=[fig.layout.title.text for fig in imagePlots])

        for i, plot in enumerate(imagePlots):
            for trace in plot['data']:  # Add each trace to the combined plot
                combined_plot.add_trace(trace, row=i + 1, col=1)

            # Copy over the x and y axis titles from the original plot
            combined_plot.update_xaxes(title_text=plot.layout.xaxis.title.text, row=i + 1, col=1)
            combined_plot.update_yaxes(title_text=plot.layout.yaxis.title.text, row=i + 1, col=1)

        # Explicitly set the width and height in the layout
        combined_plot.update_layout(
            height=500 * len(imagePlots),  # Height adjusted based on number of plots
            width=2000,  # Set the desired width here, adjust as needed
            showlegend=True,
        )

        # Save the combined plot as a single image
        combined_plot.write_image(f'{figurePath}.{config.imageFormat}', height=500 * len(imagePlots), width=2000)

    else:
        # Combine all figures into a grid when nColumns > 1
        imagePlots[0].update_layout(
            height=500 * ceil(len(figureList) / config.imageColumns),
            width=700 * config.imageColumns,  # Adjust width based on column number
            showlegend=True,
        )
        imagePlots[0].write_image(f'{figurePath}.{config.imageFormat}', height=500 * ceil(len(figureList) / config.imageColumns),
                                  width=700 * config.imageColumns)  # type: ignore


def create_cursor_plots(columnNr, config, figurePath, imagePlotsCursors, ranksCursor):
    # Handle the cursor plots (which are tables)
    if len(ranksCursor) > 0:
        cursor_path = figurePath + "_cursor"
        if columnNr in (1,2,3):
            # Create a combined plot for tables using the 'table' spec type
            combined_cursor_plot = make_subplots(rows=len(imagePlotsCursors)*2, cols=1,
                                                 specs=[[{"type": "table"}]] * len(imagePlotsCursors),
                                                 # 'table' type for each subplot
                                                 subplot_titles=[fig.layout.title.text for fig in imagePlotsCursors])
            for i, cursor_plot in enumerate(imagePlotsCursors):
                for trace in cursor_plot['data']:  # Add each trace (table) to the combined cursor plot
                    combined_cursor_plot.add_trace(trace, row=i + 1, col=1)

            # Explicitly set width and height in the layout for table plots
            combined_cursor_plot.update_layout(
                height=500 * len(imagePlotsCursors),
                width=600,  # Set the desired width for tables
                showlegend=False,
            )

            # Save the combined table plot as a single image
            combined_cursor_plot.write_image(f'{cursor_path}.{config.imageFormat}', height=500 * len(imagePlotsCursors),
                                             width=600)
        else:
            imagePlotsCursors[0].update_layout(
                height=500 * ceil(len(ranksCursor) / columnNr),
                width=500 * config.imageColumns,  # Adjust width for multiple columns
                showlegend=False,
            )
            imagePlotsCursors[0].write_image(f'{cursor_path}.{config.imageFormat}',
                                             height=500 * ceil(len(ranksCursor) / columnNr),
                                             width=500 * config.imageColumns)


def setupPlotLayout(rankName, config, figureList, htmlPlots, imagePlots, rank):
    lst: List[Tuple[int, List[go.Figure]]] = []
    if config.genHTML:
        lst.append((config.htmlColumns, htmlPlots))
    if config.genImage:
        lst.append((config.imageColumns, imagePlots))

    for columnNr, plotList in lst:
        if columnNr == 1 and plotList == imagePlots or columnNr in (1,2,3) and plotList == htmlPlots:
            for fig in figureList:
                # Create a direct Figure instead of subplots when there's only 1 column
                plotList.append(go.Figure())  # Normal figure, no subplots
                plotList[-1].update_layout(
                    title=fig.title,  # Add the figure title directly
                    plot_bgcolor='#d8d8d8',
                    height=500,  # Set height for the plot
                    legend=dict(
                        orientation="h",
                        yanchor="top",
                        y=1.22,
                        xanchor="left",
                        x=0.12,
                    )
                )
        else:
            plotList.append(make_subplots(rows=ceil(len(figureList) / columnNr), cols=columnNr))
            plotList[-1].update_layout(height=500 * ceil(len(figureList) / columnNr))  # type: ignore
            if plotList == imagePlots and rankName is not None:
                plotList[-1].update_layout(title_text=rankName)  # type: ignore


def create_css(resultsDir):

    css_path = join(resultsDir, "mtb.css")
    
    css_content = r'''body {
  font-family: Arial, Helvetica, sans-serif;
}
		
.navbar {
  overflow: hidden;
  background-color: #02525e;
  font-family: Arial, Helvetica, sans-serif;
}

.navbar {
  overflow: hidden;
  background-color: #02525e;
  font-family: Arial, Helvetica, sans-serif;
}

.navbar a {
  float: left;
  font-size: 16px;
  color: white;
  text-align: center;
  padding: 14px 16px;
  text-decoration: none;
}

.dropdown {
  float: left;
  overflow: hidden;
}

.dropdown .dropbtn {
  font-size: 16px;  
  border: none;
  outline: none;
  color: white;
  padding: 14px 16px;
  background-color: inherit;
  font-family: inherit;
  margin: 0;
}

.navbar a:hover, .dropdown:hover .dropbtn {
  background-color: #ddd;
  color: black;
}

.dropdown-content {
  display: none;
  position: absolute;
  background-color: #f9f9f9;
  min-width: 160px;
  box-shadow: 0px 8px 16px 0px rgba(0,0,0,0.2);
  z-index: 1;
}

.dropdown-content a {
  float: none;
  color: black;
  padding: 12px 16px;
  text-decoration: none;
  display: block;
  text-align: left;
}

.dropdown-content a:hover {
  background-color: #ddd;
}

.dropdown:hover .dropdown-content {
  display: block;
  
td {
  height: 50px;
  vertical-align: bottom;
}'''
    
    with open(f'{css_path}', 'w') as file:
        file.write(css_content)        
        
        
def create_html(plots: List[go.Figure], goCursorList: List[go.Figure], path: str, rankName: str, rank: int,
                config: ReadConfig, rankList) -> None:
                
    source_list = '<div style="text-align: left; margin-top: 75px;">'
    source_list += '<h2><div id="Source">Source data:</div></h2>'
    for group in config.simDataDirs:
        source_list += f'<p>{group[0]} = <a href="file:///{abspath(group[1])}" >{abspath(group[1])}</a></p>'

    source_list += '</div>'

    html_content = create_html_plots(config.htmlColumns, plots)
    html_content_cursors = genCursorHTML(config.htmlCursorColumns, goCursorList)
    
    # Create Dropdown Content for the Navbar
    idx = 0
    dropdown_content = ''
    while idx < len(rankList):
        dropdown_content += f'<a href="{rankList[idx]}.html">Rank {rankList[idx]}</a>\n'
        idx += 5
    
    # Determine the Previous and Next Rank html page for the Navbar
    idx = rankList.index(rank)
    rankPrev = rankList[idx-1]
    rankNext = rankList[idx+1 if idx+1 < len(rankList) else 0]
    
    full_html_content = f'''<html>
  <head>
    <meta name="viewport" content="width=device-width, initial-scale=1" charset="utf-8">
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
    <link rel="stylesheet" href="mtb.css"    
  </head>
  <body>
	<div class="navbar">
	  <a href="{rankPrev}.html" > &laquo; Previous Rank</a>
	  <a href="{rankNext}.html" > Next Rank &raquo;</a>
	  <div class="dropdown">
		<button class="dropbtn">More Ranks
		  <i class="fa fa-caret-down"></i>
		</button>
		<div class="dropdown-content">
          {dropdown_content}
		</div>
	  </div> 
	</div>
    <script>
        function showHelp() {{
        alert("Use Alt+PageUp to go to the previous rank\\nAnd Alt+PageDown to go to the next rank");
        }}
        document
            .addEventListener("keydown",
                function (event) {{
                    if (event.altKey && event.key === "PageUp") {{
                        event.preventDefault();
                        window.location.href = "{rankPrev}.html";
                    }} else if (event.altKey && event.key === "PageDown") {{
                        event.preventDefault();
                        window.location.href = "{rankNext}.html";
                    }} else if (event.altKey && event.key === "h") {{
                        event.preventDefault();
                        showHelp();
                    }}
                }});
    </script>
    <script type="text/javascript" id="MathJax-script" async
      src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js">
    </script>
    <script>
    MathJax = {{
      tex: {{
        inlineMath: [['$', '$'], ['\\(', '\\)']], // Allows single dollar for inline math
        displayMath: [['$$', '$$'], ['\\[', '\\]']] // Allows double dollar for block math
      }},
      svg: {{
        fontCache: 'global'
      }}
    }};
    </script>
    <h1>Rank {rank}: {rankName}</h1>
    <h4><a href="#Figures">Figures</a>&emsp;<a href="#Cursors">Cursors</a>&emsp;<a href="#Source">Source Data</a></h4>
    <br>
    {html_content}
    {html_content_cursors}
    {source_list}
    <p><center><a href="https://github.com/Energinet-AIG/MTB" target="_blank">Generated with Energinet's Model Testbench (MTB)</a></center></p>
  </body>
</html>'''

    with open(f'{path}.html', 'w', encoding='utf-8') as file:
        file.write(full_html_content)


def create_html_plots(columns, plots):
    if columns in (1,2,3):
        figur_links = '<div style="text-align: left; margin-top: 1px;">'
        figur_links += '<h2><div id="Figures">Figures:</div></h2><br>'
        for p in plots:
            plot_title: str = p['layout']['title']['text']  # type: ignore
            plot_ref = plot_title.replace('$','') # For future use with MathJax
            figur_links += f'<a href="#{plot_ref}">{plot_title}</a>&emsp;'

        figur_links += '</div>'
    else:
        figur_links = ''

    html_content = figur_links
    html_content += '<table style="width:100%">'
    html_content += '<tr>'
    for i in range(columns):
        html_content += f'<th style"width:{round(100/columns)}%"> &nbsp; </th>'
    html_content += '<tr>'
    for i, p in enumerate(plots):
        plot_title: str = p['layout']['title']['text']  # type: ignore
        plot_ref = plot_title.replace('$','') # For future use with MathJax

        if ((i+1) % columns) == 1:
            html_content += '<tr>'
        html_content += f'<td><div id="{plot_ref}">' + p.to_html(full_html=False,
                                                                   include_plotlyjs='cdn',
                                                                   include_mathjax='cdn',
                                                                   default_width='100%') + '</div></td>'  # type: ignore
        if ((i+1) % columns) == 0:
            html_content += '</tr>'

    html_content += '</table>'
    return html_content


def main() -> None:
    config = ReadConfig()

    print('Starting plotter main thread')

    # Output config
    print('Configuration:')
    for setting in config.__dict__:
        print(f'\t{setting}: {config.__dict__[setting]}')

    print()

    resultDict = mapResultFiles(config)
    figureDict = readFigureSetup('figureSetup.csv')
    cursorDict = readCursorSetup('cursorSetup.csv')
    settingsDf = pd.read_excel(config.optionalCasesheet, sheet_name='Settings', header=0)     #Read the 'Settings' sheet 
    settingDict = dict(zip(settingsDf['Name'],settingsDf['Value']))
    caseGroup = settingDict['Casegroup']
    casesDf = pd.read_excel(config.optionalCasesheet, sheet_name=f'{caseGroup} cases', header=[0, 1])
    casesDf = casesDf.iloc[:, :60]     #Limit the DataFrame to the first 60 columns
    
    colorSchemeMap = colorMap(resultDict)
    
    if not exists(config.resultsDir):
        makedirs(config.resultsDir)

    create_css(config.resultsDir)

    threads: List[Thread] = list()

    for rank in resultDict.keys():
        if config.threads > 1:
            threads.append(Thread(target=drawPlot,
                                  args=(rank, resultDict, figureDict, casesDf, colorSchemeMap, cursorDict, settingDict, config)))
        else:
            drawPlot(rank, resultDict, figureDict, casesDf, colorSchemeMap, cursorDict, settingDict, config)

    NoT = len(threads)
    if NoT > 0:
        sched = threads.copy()
        inProg: List[Thread] = []

        while len(sched) > 0:
            for t in inProg:
                if not t.is_alive():
                    print(f'Thread {t.native_id} finished')
                    inProg.remove(t)

            while len(inProg) < config.threads and len(sched) > 0:
                nextThread = sched.pop()
                nextThread.start()
                print(f'Started thread {nextThread.native_id}')
                inProg.append(nextThread)

            time.sleep(0.5)

    print('Finished plotter main thread')


if __name__ == "__main__":
    main()

if LOG_FILE:
    LOG_FILE.close()
