# WOIDE

## A Word OMDoc IDE

This plugin provides the necessary functionality to create semantic annotations within Microsoft Office Word documents.
WOIDE is a proof of concept and was developed to be a Microsoft Office Word variant of sTeX, which is currently only
available for LATEX.

With WOIDE annotations can be created, deleted and exported to SHTML (semantic HTML), furthermore annotation tags can be
toggled between three different display types.

## Installation

After cloning this repository run:
``npm install``

To run the development server and to initialize a Microsoft Office Word instance with the respective plugin manifest
run:
``npm run start:desktop``

A webpack server should launch alongside with a Microsoft Office Word instance which has the plugin preloaded.
The plugin can be opened and used by clicking "Show Taskpane" in the ribbon menu.

## Usage

Usage of WOIDE within Microsoft Office is fairly straight forward, notable is the implementation of a document URI input
which should be used to set the base URI for the document, all further annotations will use this information as a URI
prefix. Upon entering the URI the value will be saved automatically after the input is completed.

WOIDE can be configured on an administrative side by customizing the annotation type definitions, this can be done
in ``./src/taskpane/export.js``