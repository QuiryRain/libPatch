#!/usr/bin/env python3
# -*- coding: utf8 -*-
from xlsxwriter.xmlwriter import XMLwriter
from xlsxwriter.workbook import Workbook
from xlsxwriter.packager import Packager
from xlsxwriter.worksheet import Worksheet, CellFormulaTuple
from xlsxwriter.relationships import Relationships
from xlsxwriter.contenttypes import ContentTypes
from warnings import warn


WPS_APP_DOCUMENT = 'application/vnd.wps-officedocument.'
WPS_DOCUMENT_SCHEMA = "http://www.wps.cn/officeDocument/2020"


class ContentTypesLib(ContentTypes):
    def _add_cellimages(self):
        self._add_override(
            ("/xl/cellimages.xml", WPS_APP_DOCUMENT + "cellimage+xml")
        )

class CellImages(XMLwriter):
    def __init__(self):
        """
        Constructor.

        """

        super().__init__()
        self.has_dynamic_functions = False
        self.has_embedded_images = False
        self.num_embedded_images = 0

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self, image_hashes):
        # Assemble and write the XML file.

        if self.num_embedded_images > 0:
            self.has_embedded_images = True

        # Write the XML declaration.
        self._xml_declaration()

        # Write the metadata element.
        self._write_cellimages()

        # Write the metadataTypes element.
        self._write_cellimages_sub_etc_elements(image_hashes)

        self._xml_end_tag("etc:cellImages")

        # Close the file.
        self._xml_close()

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_cellimages(self):
        attributes = [
            ("xmlns:xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"),
            ("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
            ("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
            ("xmlns:etc", "http://www.wps.cn/officeDocument/2017/etCustomData"),
        ]

        self._xml_start_tag("etc:cellImages", attributes)

    def _write_cellimages_sub_etc_elements(self, image_hashes):
        for index, hash_key in enumerate(image_hashes, start=1):
            self._xml_start_tag('etc:cellImage')
            self._xml_start_tag('xdr:pic')

            self._xml_start_tag('xdr:nvPicPr')
            cnvpv_attributes = [
                ("id", f"{index}"),
                ('name', f'ID_{hash_key}')
            ]
            self._xml_start_tag('xdr:cNvPr', cnvpv_attributes)
            self._xml_end_tag("xdr:cNvPr")
            self._xml_start_tag("xdr:cNvPicPr")
            self._xml_end_tag("xdr:cNvPicPr")
            self._xml_end_tag("xdr:nvPicPr")

            self._xml_start_tag("xdr:blipFill")
            blip_attributes = [
                ("r:embed", f"rId{index}"),
            ]
            self._xml_start_tag("a:blip", blip_attributes)
            self._xml_end_tag("a:blip")
            self._xml_start_tag("a:stretch")
            self._xml_start_tag("a:fillRect")
            self._xml_end_tag("a:fillRect")
            self._xml_end_tag("a:stretch")
            self._xml_end_tag("xdr:blipFill")

            self._xml_start_tag("xdr:spPr")
            self._xml_start_tag("a:xfrm")
            a_off_attributes = [
                ("x", "0"),
                ("y", "0"),
            ]
            self._xml_start_tag("a:off", a_off_attributes)
            self._xml_end_tag("a:off")
            a_ext_attributes = [
                ("cx", "5734050"),
                ("cy", "8105775"),
            ]
            self._xml_start_tag("a:ext", a_ext_attributes)
            self._xml_end_tag("a:ext")
            self._xml_end_tag("a:xfrm")
            a_prstGeom_attributes = [
                ("prst", "rect"),
            ]
            self._xml_start_tag("a:prstGeom", a_prstGeom_attributes)
            self._xml_start_tag("a:avLst")
            self._xml_end_tag("a:avLst")
            self._xml_end_tag("a:prstGeom")
            self._xml_end_tag("xdr:spPr")
            self._xml_end_tag("xdr:pic")
            self._xml_end_tag("etc:cellImage")


class RelationshipsLib(Relationships):
    def _add_cellimages_relationship(self, rel_type, target, target_mode=None):
        rel_type = WPS_DOCUMENT_SCHEMA + rel_type
        self.relationships.append((rel_type, target, target_mode))


class PackagerLib(Packager):

    def _create_package(self):
        filename = super()._create_package()
        self._write_cellimages_file()
        return filename

    def _write_content_types_file(self):
        # Write the ContentTypes.xml file.
        content = ContentTypesLib()
        content._add_image_types(self.workbook.image_types)

        self._get_table_count()

        worksheet_index = 1
        chartsheet_index = 1
        for worksheet in self.workbook.worksheets():
            if worksheet.is_chartsheet:
                content._add_chartsheet_name("sheet" + str(chartsheet_index))
                chartsheet_index += 1
            else:
                content._add_worksheet_name("sheet" + str(worksheet_index))
                worksheet_index += 1

        for i in range(1, self.chart_count + 1):
            content._add_chart_name("chart" + str(i))

        for i in range(1, self.drawing_count + 1):
            content._add_drawing_name("drawing" + str(i))

        if self.num_vml_files:
            content._add_vml_name()

        for i in range(1, self.table_count + 1):
            content._add_table_name("table" + str(i))

        for i in range(1, self.num_comment_files + 1):
            content._add_comment_name("comments" + str(i))

        # Add the sharedString rel if there is string data in the workbook.
        if self.workbook.str_table.count:
            content._add_shared_strings()

        # Add vbaProject (and optionally vbaProjectSignature) if present.
        if self.workbook.vba_project:
            content._add_vba_project()
            if self.workbook.vba_project_signature:
                content._add_vba_project_signature()

        # Add the custom properties if present.
        if self.workbook.custom_properties:
            content._add_custom_properties()

        # Add the metadata file if present.
        if self.workbook.has_metadata:
            content._add_metadata()

        if self.workbook.has_cellimages:
            content._add_cellimages()

        # Add the metadata file if present.
        if self.workbook._has_feature_property_bags():
            content._add_feature_bag_property()

        # Add the RichValue file if present.
        if self.workbook.embedded_images.has_images():
            content._add_rich_value()

        content._set_xml_writer(self._filename("[Content_Types].xml"))
        content._assemble_xml_file()

    def _write_workbook_rels_file(self):
        # Write the _rels/.rels xml file.
        rels = RelationshipsLib()

        worksheet_index = 1
        chartsheet_index = 1

        for worksheet in self.workbook.worksheets():
            if worksheet.is_chartsheet:
                rels._add_document_relationship(
                    "/chartsheet", "chartsheets/sheet" + str(chartsheet_index) + ".xml"
                )
                chartsheet_index += 1
            else:
                rels._add_document_relationship(
                    "/worksheet", "worksheets/sheet" + str(worksheet_index) + ".xml"
                )
                worksheet_index += 1

        rels._add_document_relationship("/theme", "theme/theme1.xml")
        rels._add_document_relationship("/styles", "styles.xml")

        # Add the sharedString rel if there is string data in the workbook.
        if self.workbook.str_table.count:
            rels._add_document_relationship("/sharedStrings", "sharedStrings.xml")

        # Add vbaProject if present.
        if self.workbook.vba_project:
            rels._add_ms_package_relationship("/vbaProject", "vbaProject.bin")

        # Add the metadata file if required.
        if self.workbook.has_metadata:
            rels._add_document_relationship("/sheetMetadata", "metadata.xml")

        # Add WPS cellimages file if required.
        if self.workbook.has_cellimages:
            rels._add_cellimages_relationship("/cellImage", "cellimages.xml")

        # Add the RichValue files if present.
        if self.workbook.embedded_images.has_images():
            rels._add_rich_value_relationship()

        # Add the checkbox/FeaturePropertyBag file if present.
        if self.workbook._has_feature_property_bags():
            rels._add_feature_bag_relationship()

        rels._set_xml_writer(self._filename("xl/_rels/workbook.xml.rels"))
        rels._assemble_xml_file()

    def _write_cellimages_file(self):
        if not self.workbook.has_cellimages:
            return

        cellimages = CellImages()
        cellimages.has_dynamic_functions = self.workbook.has_dynamic_functions
        cellimages.num_embedded_images = len(self.workbook.embedded_images.images)

        image_hashes = [_[:32] for _ in self.workbook.embedded_images.image_indexes.keys()]


        cellimages._set_xml_writer(self._filename("xl/cellimages.xml"))
        cellimages._assemble_xml_file(image_hashes)


        rels = RelationshipsLib()
        for image_index in range(1, cellimages.num_embedded_images+1):
            image_type = self.workbook.embedded_images.images[image_index - 1].image_type.lower()
            rels._add_document_relationship("/image", f'media/image{image_index}.{image_type}')
        rels._set_xml_writer(self._filename("xl/_rels/cellimages.xml.rels"))
        rels._assemble_xml_file()


class WorksheetLib(Worksheet):
    def wps_embed_image(self, row, col, source, options=None):
        # Check insert (row, col) without storing.
        if self._check_dimensions(row, col):
            warn(f"Cannot embed image at ({row}, {col}).")
            return -1

        if options is None:
            options = {}

        # Convert the source to an Image object.
        image = self._image_from_source(source, options)
        image._set_user_options(options)

        cell_format = options.get("cell_format", None)

        if image.url:
            if cell_format is None:
                cell_format = self.default_url_format

            self.ignore_write_string = True
            self.write_url(row, col, image.url, cell_format)
            self.ignore_write_string = False

        image_index = self.embedded_images.get_image_index(image)

        value = f'DISPIMG("ID_{image._digest[:32]}",1)'
        self.table[row][col] = CellFormulaTuple(f'_xlfn.{value}', cell_format, f'={value}')

        return 0


class WorkbookLib(Workbook):
    worksheet_class = WorksheetLib

    def __init__(self, filename=None, options=None):
        super().__init__(filename, options)
        self.has_cellimages = False

    def _prepare_metadata(self):
        # Set the metadata rel link.
        self.has_embedded_images = self.embedded_images.has_images()
        self.has_metadata = self.has_embedded_images
        self.has_cellimages = self.has_embedded_images

        for sheet in self.worksheets():
            if sheet.has_dynamic_arrays:
                self.has_metadata = True
                self.has_cellimages = True
                self.has_dynamic_functions = True

    def _get_packager(self):
        # Get and instance of the Packager class to create the xlsx package.
        # This allows the default packager to be over-ridden.
        return PackagerLib()