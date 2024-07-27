# frozen_string_literal: true

# A Ruby Minimalist ODS
require 'rubygems'
require 'stringio'
require 'zip'
require 'date'

class MinimalistODS
  MIMETYPE = 'application/vnd.oasis.opendocument.spreadsheet'
  META_TEMPLATE = <<~XML
    <?xml version="1.0" encoding="UTF-8"?>
    <office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                          xmlns:xlink="http://www.w3.org/1999/xlink"
                          xmlns:dc="http://purl.org/dc/elements/1.1/"
                          xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
                          xmlns:ooo="http://openoffice.org/2004/office"
                          xmlns:grddl="http://www.w3.org/2003/g/data-view#"
                          grddl:transformation="http://docs.oasis-open.org/office/1.2/xslt/odf2rdf.xsl"
                          office:version="1.2">
      <office:meta>
        <meta:generator>ARMO</meta:generator>
        <meta:initial-creator>:CREATOR</meta:initial-creator>
        <dc:creator>:CREATOR</dc:creator>
        <meta:creation-date>:TIME</meta:creation-date>
        <dc:date>:TIME</dc:date>
        <meta:editing-cycles>1</meta:editing-cycles>
      </office:meta>
    </office:document-meta>
  XML

  MANIFEST_TEMPLATE = <<~XML
    <?xml version="1.0" encoding="UTF-8"?>
    <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
      <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.spreadsheet" manifest:version="1.2" manifest:full-path="/"/>
      <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
      <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
    </manifest:manifest>
  XML

  CONTENT_HEADER = <<~XML
    <?xml version="1.0" encoding="UTF-8"?>
    <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2">
      <office:scripts/>
      <office:font-face-decls>
        <style:font-face style:name="Liberation Sans" svg:font-family="&apos;Liberation Sans&apos;" style:font-family-generic="swiss" style:font-pitch="variable"/>
        <style:font-face style:name="Mangal" svg:font-family="Mangal" style:font-family-generic="system" style:font-pitch="variable"/>
        <style:font-face style:name="Microsoft YaHei" svg:font-family="&apos;Microsoft YaHei&apos;" style:font-family-generic="system" style:font-pitch="variable"/>
        <style:font-face style:name="Segoe UI" svg:font-family="&apos;Segoe UI&apos;" style:font-family-generic="system" style:font-pitch="variable"/>
        <style:font-face style:name="Tahoma" svg:font-family="Tahoma" style:font-family-generic="system" style:font-pitch="variable"/>
      </office:font-face-decls>
      <office:automatic-styles>
        <style:style style:name="co1" style:family="table-column">
          <style:table-column-properties fo:break-before="auto" style:column-width="22.58mm"/>
        </style:style>
        <style:style style:name="ro1" style:family="table-row">
          <style:table-row-properties style:row-height="4.52mm" fo:break-before="auto" style:use-optimal-row-height="true"/>
        </style:style>
        <style:style style:name="ta1" style:family="table" style:master-page-name="Default">
          <style:table-properties table:display="true" style:writing-mode="lr-tb"/>
        </style:style>
      </office:automatic-styles>
      <office:body>
        <office:spreadsheet>
          <table:calculation-settings table:automatic-find-labels="false"/>
  XML

  CONTENT_FOOTER = <<~XML
          <table:named-expressions/>
        </office:spreadsheet>
      </office:body>
    </office:document-content>
  XML

  TABLE_TEMPLATE = <<~XML
    <table:table table:name=":NAME" table:style-name="ta1">
      <table:table-column table:style-name="co1" table:number-columns-repeated=":COL_NUMBER" table:default-cell-style-name="Default"/>
  XML

  ROW_TEMPLATE = <<~XML
    <table:table-row table:style-name="ro1">
      :CELLS
    </table:table-row>
  XML

  NUMERIC_CELL_TEMPLATE = <<~XML
    <table:table-cell office:value-type="float" office:value=":VALUE">
      <text:p>:VALUE</text:p>
    </table:table-cell>
  XML

  TEXT_CELL_TEMPLATE = <<~XML
    <table:table-cell office:value-type="string" calcext:value-type="string">
      <text:p>:VALUE</text:p>
    </table:table-cell>
  XML

  TABLE_OPEN = 0
  TABLE_CLOSED = 1
  FILE_CLOSED = 2

  class MinimalistOODSError < StandardError
  end

  class TableAlreadyOpened < MinimalistOODSError
    def initialize
      super('The last table is still opened')
    end
  end

  class InvalidRowLength < MinimalistOODSError
    def initialize(expected, got)
      super("The number of rows doesn't match. Expected: #{@expected}, got: #{got}")
    end
  end

  class TableNotOpened < MinimalistOODSError
    def initialize
      super('Currently, there is not table opened')
    end
  end

  class FileClosed < MinimalistOODSError
    def initialize
      super('The file is now closed')
    end
  end

  class FileNotClosed < MinimalistOODSError
    def initialize
      super('The file is still opened')
    end
  end

  class InvalidParameter < MinimalistOODSError
  end

  attr_reader :zip_buffer, :content_buffer, :save_as, :save_to_disk, :creator

  def initialize(save_as = nil, creator = 'minimalist-ods', **options)
    @save_as = save_as
    @creator = creator

    @save_to_disk = options.fetch(:save_to_disk, true)
    raise InvalidParameter, 'Filename is required if save to disk is enabled' if save_as.nil? && save_to_disk

    init_content
  end

  def open_table(table_name, cols_number)
    raise FileClosed if @status == FILE_CLOSED
    raise TableAlreadyOpened if @status == TABLE_OPEN
    raise InvalidParameter, "Got invalid value `#{cols_number}' for table size" if cols_number.nil? || !cols_number.is_a?(Integer) || !cols_number.positive?

    @cols_number = cols_number

    table_header = TABLE_TEMPLATE.gsub(':NAME', table_name).gsub(':COL_NUMBER', cols_number.to_s)
    content_buffer.write(table_header)
    @status = TABLE_OPEN
  end

  def add_row(row)
    raise InvalidRowLength.new(@cols_number, row.size) if row.size != @cols_number

    cells = row.map { |cell| cell_to_xml(cell) }.join
    content_buffer.write(ROW_TEMPLATE.gsub(':CELLS', cells))
  end

  def close_table
    raise TableNotOpened if @status == TABLE_CLOSED

    content_buffer.write('</table:table>')
    @status = TABLE_CLOSED
  end

  def close_file
    raise TableAlreadyOpened if @status == TABLE_OPEN

    content_buffer.write(CONTENT_FOOTER)
    @status = FILE_CLOSED
    create_zip
    write_file! if save_to_disk
  end

  def file_buffer
    raise FileNotClosed unless @status == FILE_CLOSED

    zip_buffer.string
  end

  private

  def init_content
    @content_buffer = StringIO.new
    content_buffer.write(CONTENT_HEADER)
    @status = TABLE_CLOSED
  end

  def format_meta
    META_TEMPLATE.gsub(':CREATOR', creator).gsub(':TIME', Time.now.strftime('%Y-%m-%dT%H:%M:%S'))
  end

  def create_zip
    @zip_buffer = Zip::OutputStream.write_buffer do |zip_handler|
      zip_handler.put_next_entry('mimetype')
      zip_handler.write(MIMETYPE)

      zip_handler.put_next_entry('meta.xml')
      zip_handler.write(format_meta)

      zip_handler.put_next_entry('META-INF/manifest.xml')
      zip_handler.write(MANIFEST_TEMPLATE)

      zip_handler.put_next_entry('content.xml')
      zip_handler.write(content_buffer.string)
    end
  end

  def write_file!
    File.open(save_as, 'wb') do |file|
      file.write(get_file_buffer)
    end
  end

  def cell_to_xml(cell)
    if numeric?(cell.to_s)
      NUMERIC_CELL_TEMPLATE.gsub(':VALUE', cell.to_s)
    else
      TEXT_CELL_TEMPLATE.gsub(':VALUE', normalize(cell.to_s))
    end
  end

  def numeric?(str)
    true if Float(str)
  rescue StandardError
    false
  end

  def normalize(str)
    str.gsub(/[\x00-\x09\x0B\x0C\x0E-\x1F\x7F]/, '').gsub(/[&<>]/, '&' => '&amp;', '<' => '&lt;', '>' => '&gt;').encode('UTF-8')
  end
end
