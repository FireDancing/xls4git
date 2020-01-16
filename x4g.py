#!/usr/bin/env python3

# === xls4git (x4g) ===
# xls4git: Enables source-code revisioning of XML-based Microsoft Excel files
# (c) 2020 by FireDancing
#
# x4g.py: high-level Python script to control the export and building.

# Contribution: Most useful links (amongst others) used to implement x4g
# https://www.ozgrid.com/forum/forum/help-forums/excel-general/67058-import-modules-userforms-from-workbook
# https://stackoverflow.com/questions/17694393/executing-an-excel-macro-from-python
# https://gist.github.com/dreikanter/5650973

import argparse
import configparser
import os
import shutil
import zipfile
import sys
import win32com.client as win32

import xml.dom.minidom

def existsPath(path, required=False):
    # returns true, if the path exists. In case required is set the path does not exist an exception is raised
    try:
        if os.path.exists(path):
            return True
        else:
            if required:
                print('Path "' + path + '" not found!')
                raise
            else:
                return False
    except:
        raise

def getConfig(config, section, option):
    # returns a specific configuration option
    try:
        return config.get(section, option)
    except configparser.NoSectionError:
        print('Unknown section: "' + section + '"')
        raise
    except configparser.NoOptionError:
        print('Unknown option: "' + option + '"')
        raise

def copyFolderStructure(source, dest):
    # copies a folder structure overwriting existing files
    # from: https://gist.github.com/dreikanter/5650973
    for root, dirs, files in os.walk(source):
        if not os.path.isdir(root):
            os.makedirs(root)

        for file in files:
            rel_path = root.replace(source, '').lstrip(os.sep)
            dest_path = os.path.join(dest, rel_path)

            if not os.path.isdir(dest_path):
                os.makedirs(dest_path)

            shutil.copyfile(os.path.join(root, file), os.path.join(dest_path, file))


def exportXLS(config_args, config):
    # exports the Excel sources to the particular folders
    
    # checking configuration
    try:
        print('   - reading configuration from "' + config_args + '"')
        # XLSBase
        xls_base_folder = getConfig(config, 'XLSBase', 'XLSBaseFolder')
        xls_base_file = getConfig(config, 'XLSBase', 'XLSBaseFile')
        xls_base = xls_base_folder + xls_base_file
        existsPath(xls_base, True)
    
        # XLSSource
        xls_source_folder = getConfig(config, 'XLSSource', 'XLSSourceFolder')
        existsPath(xls_source_folder, True)
        xls_source = xls_source_folder + xls_base_file
        
        # VBA-sources
        xls_source_vba_folder = getConfig(config, 'XLSSource', 'XLSSourceVBAFolder')
        xls_source_vba_path = xls_source_folder + xls_source_vba_folder
        existsPath(xls_source_vba_path, True)
            
        # XLS-sources
        xls_source_xml_folder = getConfig(config, 'XLSSource', 'XLSSourceXMLFolder')
        xls_source_xml_path = xls_source_folder + xls_source_xml_folder
        existsPath(xls_source_xml_path, True)
        
        # VBA-Export handler
        xls_vba_handler = getConfig(config, 'XLSSource', 'XLSSourceVBAHandler')
        existsPath(xls_vba_handler, True)
        
        # VBA-PostTreament
        xls_vba_posttreatment_filetype = getConfig(config,  'XLSSource', 'XLSSourceVBAPostTreatmentFiletypes')
        xls_vba_posttreatment_toplinesdelete = getConfig(config,  'XLSSource', 'XLSSourceVBAPostTreatmentTopLinesDelete')
        
        # VBA-Binary
        xls_xml_binary_folder = getConfig(config, 'XLSExport', 'XLSExportVBAFolder')
        xls_xml_binary_file = getConfig(config, 'XLSExport', 'XLSExportVBABinary')
        xls_xml_binary_path = xls_source_xml_path + xls_xml_binary_folder
        xls_xml_binary_file_path = xls_xml_binary_path + xls_xml_binary_file
        xls_xml_binary_replace = getConfig(config, 'XLSExport', 'XLSExportXMLContentTypesBinaryReference')
        
        # Content_Types
        xls_xml_content_types_file = getConfig(config, 'XLSExport', 'XLSExportXMLContentTypes')
        xls_xml_content_types_path = xls_source_xml_path + xls_xml_content_types_file
        print('         --> done')

        # copy the Excel-File to the source-folder
        print('   - creating temporary copy of XLS-file "' + xls_base_file + '" from "' + xls_base_folder + '" at "' + xls_source_folder + '"')
        shutil.copy(xls_base, xls_source_folder)
        print('         --> done')
        
        # Clearing the VBA-folder
        print('   - purging VBA-folder "' + xls_source_vba_path + '"')
        for root, dirs, files in os.walk(xls_source_vba_path):
            for f in files:
                os.unlink(os.path.join(root, f))
            for d in dirs:
                shutil.rmtree(os.path.join(root, d))
        print('         --> done')
                
        # Clearing the XML-folder
        print('   - purging XML-folder "' + xls_source_xml_path + '"')
        for root, dirs, files in os.walk(xls_source_xml_path):
            for f in files:
                os.unlink(os.path.join(root, f))
            for d in dirs:
                shutil.rmtree(os.path.join(root, d))
        print('         --> done')
                
        # Extracting VBA-sources
        #   done indirectly via separate Excel-macro
        print('   - exporting VBA-code from temporary copy of XLS-file "' + xls_base_file + '" using "' + xls_vba_handler + '" to "' + xls_source_vba_path + '"')        
        excel_app = win32.Dispatch('Excel.Application')
        x4g_workbook = excel_app.Workbooks.Open(xls_vba_handler)
        x4g_worksheet = x4g_workbook.Worksheets('x4g_VBAHandler')
        x4g_worksheet.Cells(6,3).Value = xls_source.replace('/', '\\')
        x4g_worksheet.Cells(7,3).Value = xls_source_vba_path.replace('/', '\\')
        x4g_workbook.Save()
        excel_app.Application.Run("exportVBACode")
        x4g_workbook.Close()
        excel_app.Application.Quit()
        print('         --> done')
        
        # VBA PostTreament
        print('   - removing the first ' + xls_vba_posttreatment_toplinesdelete + ' lines in all "' + xls_vba_posttreatment_filetype + '"-files in folder "' + xls_source_vba_path + '"')        
        for root, dirs, files in  os.walk(xls_source_vba_path):
            for file in files:
                file_name, file_extension = os.path.splitext(file)
                if file_extension == xls_vba_posttreatment_filetype:
                    pt_file = open(xls_source_vba_path + file, 'r')
                    lines = [line for line in pt_file.readlines()]
                    pt_file.close()

                    pt_file = open(xls_source_vba_path + file, 'w')
                    for line in lines[int(xls_vba_posttreatment_toplinesdelete):]:
                        pt_file.write(line)
                    pt_file.close()
        print('         --> done')
        
        # Extracting XML-sources
        print('   - extracting temporary copy of XLS-file "' + xls_base_file + '" to "' + xls_source_xml_path + '"')
        zipfile.ZipFile(xls_source, 'r').extractall(xls_source_xml_path)
        print('         --> done')

        # Removing binary VBAProject file
        print('   - removing binary VBA-project file "' + xls_xml_binary_file_path + '"')
        existsPath(xls_xml_binary_path, True)
        existsPath(xls_xml_binary_file_path, True)
        os.remove(xls_xml_binary_file_path)
        print('         --> done')
            
        # Adapting content_types-file
        print('   - removing binary references from content-types file "' + xls_xml_content_types_path + '"')
        existsPath(xls_xml_content_types_path, True)
        xml_file = open(xls_xml_content_types_path, 'r')
        xml_string = xml_file.read()
        xml_file.close()        
        xml_string = xml_string.replace(xls_xml_binary_replace, '')
        xml_file = open(xls_xml_content_types_path, 'w')
        xml_file.write(xml_string)
        xml_file.close()
        print('         --> done')

        # build nice XML-files
        print('   - formatting all XML-files in path "' + xls_source_xml_path +'"')        
        for root, dirs, files in os.walk(xls_source_xml_path):
            for file in files:
                file_name, file_extension = os.path.splitext(file)
                if file_extension == ".xml":
                    file_path = root + "/" + file
                    file_path = file_path.replace('\\', '/')
                    xml_file = open(file_path, 'r', encoding="utf8")
                    xml_content = xml.dom.minidom.parseString(xml_file.read())
                    xml_file.close()
                    
                    xml_pretty_content = xml_content.toprettyxml()
                    
                    xml_file = open(file_path, 'w', encoding="utf8")
                    xml_file.write(xml_pretty_content)
                    xml_file.close()
        print('         --> done')

        # Removing temporary XLS file
        print('   - removing temporary XLS-file "' + xls_base_file + '" at "' + xls_source_folder + '"')
        os.remove(xls_source)
        print('         --> done')
        
        # finished
        print()
        print('     XLS sources successfully exported!')
        
    except Exception as e:
        raise


def buildXLS(config_args, config):
    # builds the XLS from the sources
    
    # checking configuration
    try:
        print('   - reading configuration from "' + config_args + '"')
        # XLSSource
        xls_source_folder = getConfig(config, 'XLSSource', 'XLSSourceFolder')
        existsPath(xls_source_folder, True)
        
        # VBA-sources
        xls_source_vba_folder = getConfig(config, 'XLSSource', 'XLSSourceVBAFolder')
        xls_source_vba_path = xls_source_folder + xls_source_vba_folder
        existsPath(xls_source_vba_path, True)
            
        # VBA-Export handler
        xls_vba_handler = getConfig(config, 'XLSSource', 'XLSSourceVBAHandler')
        existsPath(xls_vba_handler, True)
            
        # XML-sources
        xls_source_xml_folder = getConfig(config, 'XLSSource', 'XLSSourceXMLFolder')
        xls_source_xml_path = xls_source_folder + xls_source_xml_folder
        existsPath(xls_source_xml_path, True)
        
        # XMLBuild
        xls_build_folder = getConfig(config, 'XLSBuild', 'XLSBuildFolder')
        existsPath(xls_build_folder, True)
        xls_build_file = getConfig(config, 'XLSBuild', 'XLSBuildFile')
        xls_build = xls_build_folder + xls_build_file
        print('         --> done')

        # building Excel-file
        print('   - building XLS-file "' + xls_build + '" from XML-files in "' + xls_source_xml_path +'"')        
        xlsx_file = zipfile.ZipFile(xls_build, 'w', zipfile.ZIP_STORED)
        for root, dirs, files in  os.walk(xls_source_xml_path):
            root_path = root[len(xls_source_xml_path):]
            for file in files:
                # the tempoaray excel-file shall not be packed
                if file != xls_build_file:
                    xlsx_file.write(os.path.join(root, file), os.path.join(root_path, file))
        xlsx_file.close()
        print('         --> done')
        
      
        # adding VBA-sources
        #   done indirectly via separate Excel-macro
        print('   - adding VBA-code from folder "' + xls_source_vba_path + '" using "' + xls_vba_handler + '"')        
        excel_app = win32.Dispatch('Excel.Application')
        x4g_workbook = excel_app.Workbooks.Open(xls_vba_handler)
        x4g_worksheet = x4g_workbook.Worksheets('x4g_VBAHandler')
        x4g_worksheet.Cells(10,3).Value = xls_build.replace('/', '\\')
        x4g_worksheet.Cells(11,3).Value = xls_source_vba_path.replace('/', '\\')
        x4g_workbook.Save()
        excel_app.Application.Run("importVBACode")
        x4g_workbook.Close()
        excel_app.Application.Quit()
        print('         --> done')
        
        # finished
        print()
        print('     Excel sheet "'  + xls_build + '" successfully created!')   
        
    except Exception as e:
        raise
    
    
if __name__ == '__main__':
    
    # Read the arguments
    arg_parser = argparse.ArgumentParser(description='x4g: exporting, building and releasing the source-code of an Excel sheet ')
    arg_parser.add_argument('-c', '--config', required=True, help='The name of the x4g configuration file (ini-format)')
    arg_parser.add_argument('-a', '--action', required=True, choices=['export', 'build'], help='The type of action to be done: "export" to export the Excel-sources, "build" to build the Excel-file from the source')    
    args = arg_parser.parse_args()

    try:
        existsPath(args.config, True)
    except Exception as e:
        print('Exception: '+ type(e).__name__ + ': ' +  e.args)
        quit()

    # Read the configuration
    config = configparser.ConfigParser()
    config.read(args.config)
    
    # the further action
    try:
        if args.action == 'export':
            print('Action: Exporting XLS sources')
            exportXLS(args.config, config)
            
        if args.action == 'build':
            print('Action: Building XLS from sources')
            buildXLS(args.config, config)
    
    except Exception as e:
        print('Exception: '+ type(e).__name__ + ': ' +  e.args)
        quit()
        