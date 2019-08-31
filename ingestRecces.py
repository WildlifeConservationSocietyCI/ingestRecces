# ingestRecces.py: script to iterate over directory of excel files and for each table found therein
# create a point feature class, then dissolve those points to lines splitting by date and sorting by time,
# then append the result to a combined feature class.
# 0.1 Kim Fisher 13.03
# 0.2 Kim Fisher 14.04
# Note I'd like to do intermediate steps in_memory, but 1) Project tool doesn't allow it;
# 2) Feature Class to Feature Class reads in_memory as 'folder' and so forces shapefile output, which doesn't work
# So until 2) is allowed this isn't worth it

# TODO: validate date and time fields record by record?
# TODO: implement per-record validation based on maximum distance allowed between points?
# TODO: validate record for being GPS anomaly (sharp detours)?

# TODO: deal with csv time string conversion issue noted below
# TODO: support output to shapefile? (Need to deal with timestamps, handle long field names. Or, just move from temp gdb and then delete gdb.)
# TODO: test with different input types (xls, xlsx, csv; different time/date formats; different projection combinations)
# TODO: determine whether/how to summarize stats
# TODO: determine whether to compare dates of line FC records to those of outputFC records before appending (to avoid duplicates)

import arcpy, os
from arcpy import env

inputWS = arcpy.GetParameterAsText(0)       # D:\_data\Documents\Africa\Congo\Goualougo\input_recce_data
recceY = arcpy.GetParameterAsText(1)        # LAT
recceX = arcpy.GetParameterAsText(2)        # LONG
recceDate = arcpy.GetParameterAsText(3)     # DATE
recceTime = arcpy.GetParameterAsText(4)     # TIME
outputWS = arcpy.GetParameterAsText(5)      # D:\_data\Documents\Africa\Congo\Goualougo\output_features\recces_2004.gdb
outputFC = arcpy.GetParameterAsText(6).replace(' ', '')  # 2004 recces
inputPrj = arcpy.GetParameterAsText(7)      # GCS_WGS_1984
outputPrj = arcpy.GetParameterAsText(8)     # UTM 33N WGS 84

# If outputPrj is not a projected coordinate system, do not proceed
outputPrjInstance = arcpy.SpatialReference()
outputPrjInstance.loadFromString(outputPrj)
if outputPrjInstance.type != 'Projected':
    arcpy.AddError('Specified output coordinate system is not projected. A projected coordinate system is necessary to measure planimetric lengths and distances.')
    raise arcpy.ExecuteError

env.workspace = inputWS # set for ListFiles()
for recceFile in arcpy.ListFiles():
    # Note ArcGIS requires Office 2007 drivers in order to connect to xlsx files: http://forums.arcgis.com/threads/82779-ARc-for-Desktop-10.1-and-Microsoft-Office-2013
    # These can be installed from: http://www.microsoft.com/en-us/download/confirmation.aspx?id=23734
    arcpy.AddMessage('Working on file: %s' % recceFile)
    if recceFile[-3:] == 'xls' or recceFile[-4:] == 'xlsx': env.workspace = os.path.join(inputWS, recceFile) # set for ListTables() if file is excel format
    env.outputCoordinateSystem = inputPrj


    for recceTable in arcpy.ListTables(): # for each table in the current excel file, or each table in the current workspace generally
        # excel table names ending in _ are filters or named ranges: http://stackoverflow.com/questions/4510917/referencing-excel-sheets-with-jet-driver-sheets-are-duplicated-with-underscores
        if recceTable[-1] != '_':
            arcpy.AddMessage('Ingesting data in table %s' % recceTable)
            recceTableBasename = os.path.splitext(recceTable)[0].strip('$') # get name without file extension (csv) or $ (xls worksheet)
            pointsLyr = recceTableBasename + '_lyr'
            pointsFC = recceTableBasename + '_fc'
            if recceFile[-3:] == 'xls' or recceFile[-4:] == 'xlsx': env.workspace = os.path.join(inputWS, recceFile) # set for ListFields() since it gets reset further down in loop
            env.outputCoordinateSystem = inputPrj

            recceTableFields = arcpy.ListFields(recceTable)
            if any(recceX in s.name for s in recceTableFields) and \
                    any(recceY in s.name for s in recceTableFields) and \
                    any(recceDate in s.name for s in recceTableFields) and \
                    any(recceTime in s.name for s in recceTableFields): # make sure table has required fields

                # TODO: Clean the data. Iterate over each row of table, and only copy to cleanedRecceTable if:
                # - coordinates all 'look right': numeric; could also test for difference from previous record's (but only if previous record is for same date)
                # Test: GTAP_Recce_Data_K_Fisher.xlsx:Sheet1$:15610-15611
                # - year is a 4-digit integer
                # Anything else? Then use cleanedRecceTable to proceed.

                # Convert table to spatial point data
                arcpy.MakeXYEventLayer_management(recceTable, recceX, recceY, pointsLyr)
                # Note this has to output to a gdb to deal with 24-hr times correctly. If we need to support shp output, we'll need to set up an intermediate gdb.
                # http://forums.esri.com/Thread.asp?t=106423&c=93&f=1149
                arcpy.FeatureClassToFeatureClass_conversion(pointsLyr, outputWS, pointsFC)

                # More fun with time: sometimes arc won't be able to make recceTime in pointsFC a Date field (if recceTable is csv and clock is 24-hr, e.g.), so it will be Text.
                # For the PointsToLine sort to work, we need Date. We'll test for whether recceTime is a String type, and if it is, attempt to convert from
                # 1) 24 hr and 2) 12 hr clocks, and if neither work, delete pointsFC and bail out of loop with a warning.
                # This does not work currently, and crashes ArcGIS. When attempted in interactive python window, expression works but produces only 1/1/2001.
                recceTimeConverted = recceTime
                # for f in recceTableFields:
                #     if f.name == recceTime and f.type == 'String':
                #         recceTimeConverted = recceTime + '_'
                #         try:
                #             arcpy.AddMessage('trying 1033;H:mm:ss;;')
                #             arcpy.ConvertTimeField_management(pointsFC, recceTime, "1033;H:mm:ss;;", recceTimeConverted) # try converting from 24-hr
                #         except:
                #             try:
                #                 arcpy.AddMessage('trying 1033;h:mm:ss tt;;')
                #                 arcpy.ConvertTimeField_management(pointsFC, recceTime, "1033;h:mm:ss tt;AM;PM", recceTimeConverted) # try converting from 12-hr
                #             except: # neither time format worked
                #                 arcpy.Delete_management(pointsFC)
                #                 arcpy.AddWarning('Unable to convert %s string field holding times to date field. Skipping this table.' % recceTime)
                #                 break # break out of for f in recceTableFields, which will cause block after else not to execute
                #
                # else:
                env.workspace = outputWS
                env.outputCoordinateSystem = outputPrj
                # If inputPrj is not the same as outputPrj, project points FC to output prj so lines will have meaningful distance units
                # Initial outputPrj validation ensures outputPrj will be a projected coordinate system, allowing for length/distance measurement
                pointsFCPrj = pointsFC
                if inputPrj != outputPrj:
                    pointsFCPrj = pointsFC + 'prj'
                    arcpy.Project_management(pointsFC, pointsFCPrj, outputPrj)

                # Here's where to put per-record validation, using cleaned and projected pointsFC
                # TODO: measure distance between points. If distance / time difference > 100 m / 30 s then return error with which points

                # Turn points into lines, using date to distinguish separate lines and time to draw vertices in right order
                linesFC = pointsFC + '_line'
                arcpy.PointsToLine_management(pointsFCPrj, linesFC, recceDate, recceTimeConverted)

                # Calculate lengths of lines, to store total distance per day
                # (For gdb output we'll have Shape_Length, but below will support shapefile output, and Shape_Length is unrealistically precise)
                arcpy.AddField_management(linesFC, 'distance', 'LONG')
                arcpy.CalculateField_management(linesFC, 'distance', '!shape.length!', 'PYTHON') # length is in meters

                # Merge into main output FC
                if arcpy.Exists(outputFC):
                    arcpy.Append_management([linesFC], outputFC)
                else:
                    arcpy.CopyFeatures_management(linesFC, outputFC)

                # Delete intermediate files
                arcpy.Delete_management(pointsFC)
                if inputPrj != outputPrj: arcpy.Delete_management(pointsFCPrj)
                arcpy.Delete_management(linesFC)

            else:
                arcpy.AddWarning('Table %s is missing one of the following required specified fields: %s. Skipping this table.' % (
                recceTable, ', '.join([recceX, recceY, recceDate, recceTime])))

arcpy.AddMessage('Finished.')
