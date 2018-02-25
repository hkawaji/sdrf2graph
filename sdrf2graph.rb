#!/usr/bin/env ruby


#= Graph (investigation design graph) genertor from SDRF sheets in MAGE-tab
#
#==SYNOPSIS
#
#   ./sdrf2graph.rb --infile <infile> [options]
#
#
#==REQUIREMENTS
#
# - ruby
#   http://www.ruby-lang.org/
#
# - rexml
#   http://www.germane-software.com/software/rexml/index.html
#
# - rubyzip
#   http://raa.ruby-lang.org/project/rubyzip/
#   (% gem install rubyzip)
#
# - GraphViz
#   http://www.graphviz.org/
#   (required only when producing image files, sush as PNG or SVG
#
#
#
#==DESCRIPTION
#
#   sdrf2graph.rb is an application to produce graphical image
#   of  investigation design graph (IDG) based on SDRFs witten
#   in a MAGE-tab formatted spreadsheet(*.xlsx).
#
#   MAGE-tab is a generic framework which can represent a complicated
#   experiments and can be used for data exchange not just database
#   submission. sdrf2graph.rb aims to facilitate 1) representation
#   of complicated or large-scale experiments, and 2) exchange of
#   essential information (not just data submission). To aim this goal,
#   sdrf2graph.rb have these features:
#
#   - take *.xlsx formated spreadsheet as input,
#     (*.xlsx can be created by both of MicrosoftOffice Excel and OpenOffice.org Calc)
#
#   - a subgraph can be generated with specifying sheets used in the graph,
#     according to their names
#
#   - work as a web application so that an end user (incl. wet scientist)
#     does not need to install this application
#
#   - show protocol (edge) information explicitly to help understanding
#     of experiments in detail
#
#   - show characteristic and parameter values for node and edge, respectively
#
#   - Note: gap ('>') is not treated properly. Please contert your SDRF into
#     different sheets to represent your data without gap.
#
#   The graphical didplay is based on GraphViz (http://www.graphviz.org),
#   and this software can produce GraphViz native (dot) language or GraphViz
#   supproted image format. Note that URL can be treatd only in SVG format.
#
#
#   [Note] Tab2MAGE (http://tab2mage.sourceforge.net/) is another software
#          package to handle MAGE-tab formated files, which is written by the
#          ArrayExpress curation team. "expt_check.pl" in the package can produce
#          similar graph, as well as data validation with focusing on data
#          submission to ArrayExpress. Tab2MAGE is strongly recommended if you
#          are now submitting your data to ArrayExpress.
#
#
#   sdrf2graph.rb produces a graph according to this rule:
#
#      Each node in the investigation design graph must
#      be represented by an appropriate 'Name' column, 
#      with the graph edges given as 'Protocol REF' columns.
#      (MAGE-tab specification v1.0 March 20, 2007, Section 2.3.1)
#
#   That is, all '* Name' columns are treated as nodes, and 'Protocol REF'
#   are treated as edges. In addition, '*File' and '*Data' are also treated as nodes.
#
#   Hyperlink is supported in DOT/SVG foramt as: if a column
#   which name includes "URI", "URL", or "FTP" appears right to 'Name'
#   or 'Protocol REF', the URI is attached to the corresponding nodes
#   or edges.
# 
# 
# 
#   [MAGE-tab]
#   A simple spreadsheet-based, MIAME-supportive format for microarray data: MAGE-TAB.
#   Tim F. Rayner et al. (2006) BMC Bioinformatics, 7, 489.
#   http://www.biomedcentral.com/1471-2105/7/489
#
#   [Tab2MAGE]
#   http://tab2mage.sourceforge.net/
#
#   [Office Open XML File Formats]
#   http://www.ecma-international.org/publications/standards/Ecma-376.htm
#
#
#==OPTIONS
#
#  --infile <file>
#     input file XLSX format. REQUIRED if --server is not setted.
#     (Microsoft Office 2007 or OpenOffice.org can write a file with this format)
#
#  --outfile <file>
#     output file.
#     default: "[infile].[suffix]"
#
#  --format <format>
#      "dot"(default), "png", "svg", ...
#      default: dot
# 
#
#  --sheet_name <substr>
#      sheets which name include <substr> are used in the graph
#      default: 'sdrf'
#
#  --layout <top-to-down|left-to-right>
#      default: top-to-down
#
#  --edge_label <TRUE|FALSE>
#      default: TRUE
#
#  --server
#      run as a web application
#      default URL: http://localhost:10080/
#
#  --port
#      specify the port used by the web server
#      default: 10080
#
#  --bind_address
#      specify the IP address or hostname of the web server
#
#
#---
#Author:: Hideya  Kawaji  <http://genome.gsc.riken.jp/osc/members/Hideya_Kawaji.html>
#Version:: 1.0
#Copyright:: (c)RIKEN Omics Science Center
#License:: Ruby's
#---

# added for fantom45
require 'rdoc/ri/ri_paths'
require 'rubygems'

require 'rexml/document'
require 'zip/zipfilesystem'
require 'ftools'
require 'webrick'
require 'getoptlong'
require 'open3'
require 'rdoc/usage'

#-------------------------------------------------------
# read *.xlsx file and return a graph in "dot" language
#-------------------------------------------------------

class SdrfXlsx

  attr_accessor :config

  def initialize(conf) 
    @config = conf
  end

  # generate graph in dot language
  def to_dot
    n2n , urls , labels = self.get_name2name
    node_names = []
    node_protocols = []
    edges = []
    n2n.keys.each do |name1|
      n2n[name1].keys.each do |name2|
        if config[:edge_label] == "FALSE"
          label = n2n[name1][name2]
          edges << "\"#{name1}\" -> \"#{name2}\";"
          node_protocols << label
          node_names << name1
          node_names << name2
        elsif config[:edge_label] == "TRUE"
          label = n2n[name1][name2]
          node_names << name1
          node_names << name2
          if (label == nil) or label.match(/^\s*$/)
            edges << "\"#{name1}\" -> \"#{name2}\" ;"
          else
            url = nil
            url = urls[label] if urls.key?(label)
            id = %Q! #{label}\\n(from:#{name1.split("\\n")[0]})\\n(to:#{name2.split("\\n")[0]}) !
            urls[id] = url
            labels[id] = label
            node_protocols << id
            edges << "\"#{name1}\" -> \"#{id}\" [arrowhead = none] ;"
            edges << "\"#{id}\" -> \"#{name2}\" ;"
          end
        end
      end
    end
    out = ["digraph sample {"]
    out << "graph [rankdir = LR];" if ! config[:layout].match(/top-to-down/)
    out += node_names.uniq.collect{|x| 
      styles = []
      styles << "shape = box"
      styles << "URL=\"#{urls[x]}\"" if urls.key?(x)
      styles << "label=\"#{labels[x]}\"" if labels.key?(x)
      %! "#{x}" [#{styles.join(',')}] ; !
    }
    if config[:edge_label] != "FALSE"
      out += node_protocols.uniq.collect{|x|
        styles = []
        styles << "shape=none"
        styles << "URL=\"#{urls[x]}\"" if urls.key?(x)
        styles << "label=\"#{labels[x]}\"" if labels.key?(x)
        #label = x.sub(/\(to[^)]*\)\s*/,"").sub(/\(from[^)]*\)\s*/,"").sub(/[\\n]*\s*$/,"")
        #styles << %!label="#{label}"!
        %! "#{x}" [#{styles.join(',')}] ; !
      }
    end
    out += edges
    out << "}"
    return out
  end

  # get network (node name to node name)
  def get_name2name
    name2name = {}
    urls = {}
    labels = {}
    sdrf = self.get_spreadsheet
    sdrf[:sheets].each do |sheet_obj,sheet|
      sheet_data = get_sheet_data(sheet,sdrf[:shared_string])
      sheet_data[:rows].each do |row|
        name1 = nil
        protocol = []
        protocol_url = nil
        sheet_data[:header].each_index do |i|
          row[i] = row[i].to_s
          next if row[i].match(/^\s*$/)
          if sheet_data[:header][i].match(/Array\s*Design\s*[File|REF]/)

            labels[name1] ||= name1
            labels[name1] += "\\n(Array:#{row[i]})" if ! labels[name1].match(/\(Array:#{row[i]}/)

          elsif (!sheet_data[:header][i].match(/Protocol/)) and
             (sheet_data[:header][i].match(/Name|File|Data/))

            if sheet_data[:header][i].match(/Name/)
              row[i] = sheet_data[:header][i].sub(/\s*Name\s*/,"") + "|" + row[i]
            elsif sheet_data[:header][i].match(/File/)
              row[i] = "File|" + row[i]
            end
            if (name1 != nil) and (! name1.match(/\|\s*$/))
              name2name[name1] ||= {}
              name2name[name1][row[i]] = protocol.join("\\n")
              urls[protocol.join("\\n")] = protocol_url if protocol_url != nil
            end
            name1 = row[i]
            protocol = []
            protocol_url = nil

          elsif sheet_data[:header][i].match(/Characteristic/)

            if m = sheet_data[:header][i].match(/\[(.*)\]/)
              key = m.to_a[1]
            end
            labels[name1] ||= name1
            labels[name1] += "\\n(#{key}:" + row[i] + ")" if ! labels[name1].match(/\(#{key}:/)

          elsif sheet_data[:header][i].match(/Protocol/)

            protocol << row[i]

          elsif sheet_data[:header][i].match(/Parameter/)

            idx = protocol.length - 1
            if m = sheet_data[:header][i].match(/\[(.*)\]/)
              key = m.to_a[1]
            end
            protocol[idx] = protocol[idx] + "\\n(#{key}:" + row[i] + ")"
            urls[protocol.join("\\n")] = protocol_url if protocol_url != nil

          elsif sheet_data[:header][i].match(/URI|URL|FTP/)

            if (protocol == []) and (name1 != nil)
              urls[name1] = row[i]
              protocol_url = nil
            else
              protocol_url = row[i]
              urls[protocol.join("\\n")] = protocol_url if protocol.length > 0
            end
          end

        end
      end
    end
    return name2name , urls, labels
  end

  def get_sheet_data(sheet,shared_string)
    sheet_data ={ :header => [] , :rows => [] }
    sheet.each_index do |row_idx|
      # prepare a hash of column name => value
      col_name2value = {}
      sheet[row_idx].each do |col|
        value = col[:value]
        if col[:value_type] == "s"
          value = shared_string[value.to_i + 1] if col[:value_type] == "s"
        end
        col_name = col[:cell_name].match(/^([A-Z]+)/).to_a[1]
        col_name2value[col_name] = value
      end

      # prepare values in order of column
      col_names = col_name2value.
                  keys.
                  sort{|a,b| if a.length == b.length then a <=> b else (a.length <=> b.length) end}
      colv = []
      col_names.first.upto(col_names.last) do |n|
        if col_name2value.key?(n)
          colv << col_name2value[n]
        else
          colv << "-"
        end
      end

      # set header or values
      if row_idx == 0
        sheet_data[:header] = colv
      else
        sheet_data[:rows] << colv
      end
    end
    return sheet_data
  end

  def get_spreadsheet
    buf = {:sheets => {} , :shared_string => nil , :rId2name => nil}
    Zip::ZipFile.foreach(config[:infile]) do |zf|
      buf[:rId2name] = get_rId2name(zf) if zf.to_s.match(/xl\/workbook.xml$/)
      buf[:shared_string] = get_shared_string(zf) if zf.to_s.match(/xl\/sharedStrings\.xml$/)
      buf[:target_sheet2rId] = get_target_sheet2rId(zf) if zf.to_s.match(/xl\/_rels\/workbook.xml.rels/)
    end

    Zip::ZipFile.foreach(config[:infile]) do |zf|
      if m = zf.to_s.match(/(worksheets\/sheet.*.xml)$/) 
        sheet_id = m.to_a[1]
        if (! config.key?(:sheet_name)) or
           (config[:sheet_name] == "") or
           (config[:sheet_name] == "FALSE")
          buf[:sheets][zf] = get_sheet(zf)
        elsif (buf[:target_sheet2rId].key?(sheet_id)) &
              (buf[:rId2name].key?(buf[:target_sheet2rId][sheet_id])) &
              (buf[:rId2name][buf[:target_sheet2rId][sheet_id]].match(/#{config[:sheet_name]}/i))
          buf[:sheets][zf] = get_sheet(zf)
        end
      end
    end

    return buf
  end

  def get_sheet(zf)
    out = ""
    zf.get_input_stream do |ifh|
      doc = REXML::Document.new(ifh)
      out = _get_sheet(doc)
    end
    return out
  end

  def get_rId2name(zf)
    out = ""
    zf.get_input_stream do |ifh|
      doc = REXML::Document.new(ifh)
      out = _get_rId2name(doc)
    end
    return out
  end

  def get_shared_string(zf)
    out = ""
    zf.get_input_stream do |ifh|
      doc = REXML::Document.new(ifh)
      out = _get_shared_string(doc)
    end
    return out
  end

  def get_target_sheet2rId(zf)
    out = ""
    zf.get_input_stream do |ifh|
      doc = REXML::Document.new(ifh)
      out = _get_target_sheet2rId(doc)
    end
    return out
  end

  def _get_rId2name(doc)
    rId2name = {}
    doc.each_element("/workbook/sheets/sheet") do |t|
      rId2name[ t.attributes["r:id"] ] = t.attributes["name"]
    end
    return rId2name
  end

  def _get_shared_string(doc)
    buf = {}
    doc.each_element("/sst/si") do |si|
    si.index_in_parent
      si.each_element('.//t') do |t|
        buf[ si.index_in_parent ] ||= ""
        buf[ si.index_in_parent ]  += t.text
      end
    end
    return buf
  end

  def _get_target_sheet2rId(doc)
    target_sheet2rId = {}
    doc.each_element("/Relationships/Relationship") do |t|
      target_sheet2rId[ t.attributes["Target"] ] = t.attributes["Id"]
    end
    return target_sheet2rId
  end

  def _get_sheet(doc)
    out = []
    doc.each_element("/worksheet/sheetData/row") do |row|
      cells = []
      row.children.each do |col|
        cell_name = col.attributes["r"]
        value_type = col.attributes["t"]
        v = ""
        col.children.each do |value|
          v = value.text
        end
        cells << {:cell_name => cell_name, :value_type => value_type, :value => v}
      end
      out << cells
    end
    return out
  end

end


#-----------------------------------------
# for server
#-----------------------------------------

class SdrfGraphServer
  attr_accessor :server_config

  def initialize(conf={:port => 10080 , :bind_address => '127.0.0.1'})
    @server_config = conf
  end

  #
  # top page
  #
  class IndexServlet < WEBrick::HTTPServlet::AbstractServlet
    def do_GET(req, res)
      res['Content-Type'] = 'text/html'
      res.body = <<EOF

<body>

  <h1> SDRF2GRAPH (v1.0) </h1>

  <form action="/4/sdrf2graph/sdrf" method="post" enctype="multipart/form-data">

    <div>
    <p>
      Draw an investigation design graph
    </p>
    <p>
      <select name="layout">
      <option value="left-to-right">in left-to-right layout</option>
      <option value="top-to-down">in top-to-down layout</option>
      </select>
      <select name="edge_label">
      <option value="TRUE">with edge labels</option>
      <option value="FALSE">without edge labels</option>
      </select>
      <select name="format">
      <option value="png">as PNG (image file) file</option>
      <option value="svg">as SVG (Scalable Vector Graphics) file</option>
      <option value="dot">as DOT (GraphViz native) format file</option>
      </select>
    </p>
    <p>
      from sheets &Prime;<input type="text" name="sheet_name" value="sdrf" size="10" />&Prime; <span style="font-size: 70%"> (partial strings of the sheet name. By using the '|', multiple sheets can be specified for visualization ex. &Prime;sdrf-kd|sdrf-seq&Prime;. If no strings are specified, all sheets are used.) </span>
    </p>
    <p>
      in the spreadsheet <input type="file" name="filename" /> (*.xlsx - see <a href="/4/sdrf2graph/example.xlsx">a sample file</a>) 
      <input type="submit" "value="draw"/>
    </p>
    <br>
    <p>
      <span style="font-weight: bold">Note:</span> if the PNG does not show, please try SVG or DOT format, which normally result in smaller files.
    </p>
    </div>

  </form>

  <hr>
  <p>
    This is a web application of <a href="/4/sdrf2graph/source">sdrf2graph.rb</a>, which is independent to <a href="http://tab2mage.sourceforge.net/"> tab2mage </a>. See <a href="/4/sdrf2graph/source">here</a> for more details.
  </p>

</body>

EOF
    end
  end

  class SourceServlet < WEBrick::HTTPServlet::AbstractServlet
    def do_GET(req, res)
      res['Content-Type'] = 'text/plain'
      open($0){|ifh| res.body = ifh.read}
    end
  end

  class ExampleServlet < WEBrick::HTTPServlet::AbstractServlet
    def do_GET(req, res)
      res['Content-Type'] = 'application/octet-stream'
      open("example.xlsx"){|ifh| res.body = ifh.read}
    end
  end


  #
  # drawing part
  #
  class SdrfDrawServlet < WEBrick::HTTPServlet::AbstractServlet
    def do_POST(req, res)
      if req.query['filename']
        tmpfile = "/tmp/__sdrf2graph.#{$$}"
        open(tmpfile,"w"){|ofh| ofh.print req.query['filename']}
        config = {
          :infile => tmpfile,
          :format => "svg",
          :sheet_name => "sdrf",
          :edge_label => "TRUE" ,
          :layout => "left-to-right"
          #:layout => "top-to-down"
        }
        config.keys.each{|k| config[k] = req.query[k.to_s] if req.query[k.to_s]}

        sdrf = SdrfXlsx.new(config)
        if config[:format] == "dot"
          res.body = "<pre>" + sdrf.to_dot.join("\n") + "</pre>"
          res['Content-Type'] = 'text/html'
        else
          # exec GraphViz
          Open3.popen3("dot -T#{config[:format]}") do |stdin, stdout, stderr|
            stdin.puts sdrf.to_dot
            stdin.close
            res.body = stdout.read
          end
          if config[:format] == "svg"
            res['Content-Type'] = 'image/svg+xml'
          elsif config[:format] == "png"
            res['Content-Type'] = 'image/png'
          else
            res['Content-Type'] = 'image/unknown'
          end
        end
        File.delete tmpfile
      end
    end
  end

  #
  # set up
  #
  def run
    server = WEBrick::HTTPServer.new({
      :Port => server_config[:port],
      :BindAddress => server_config[:bind_address]}
    )
    server.mount('/sdrf', SdrfDrawServlet)
    server.mount('/', IndexServlet)
    server.mount('/source', SourceServlet)
    server.mount('/example.xlsx', ExampleServlet)
    trap('INT') { server.shutdown }
    server.start
  end

end




#-----------------------------------------------------
# command line
#-----------------------------------------------------

if __FILE__ == $0

  config = {
    :infile => nil,
    :outfile => nil,
    :format => "dot",
    :edge_label => "TRUE" ,
    :sheet_name => "sdrf" ,
    :layout => "top-to-down" ,
    :server => "FALSE" ,
    :port => 10080 ,
    :bind_address => "127.0.0.1"
  }
  opts = GetoptLong.new(
    [ '--help', '-h', GetoptLong::NO_ARGUMENT ],
    [ '--infile', '-i', GetoptLong::REQUIRED_ARGUMENT ],
    [ '--outfile', '-o', GetoptLong::OPTIONAL_ARGUMENT ],
    [ '--format', '-f', GetoptLong::OPTIONAL_ARGUMENT ],
    [ '--edge_label', '-e', GetoptLong::OPTIONAL_ARGUMENT ] ,
    [ '--sheet_name', '-s', GetoptLong::OPTIONAL_ARGUMENT ] ,
    [ '--layout', '-l',GetoptLong::REQUIRED_ARGUMENT ] ,
    [ '--server',  GetoptLong::NO_ARGUMENT ],
    [ '--port', GetoptLong::REQUIRED_ARGUMENT ] ,
    [ '--bind_address', GetoptLong::REQUIRED_ARGUMENT ]
  )
  opts.each do |opt, arg|
    if opt == "--help"
      RDoc::usage
    elsif opt == "--server"
      config[:server] = true
    else
      config[opt.sub(/^-+/,"").to_sym] = arg
    end
  end

  # start server
  if config.key?(:server) and config[:server] == true
    s = SdrfGraphServer.new(config)
    s.run
    exit
  end

  # print usage
  RDoc::usage if config[:infile] == nil

  # exec the main function
  config[:outfile] = config[:infile] + ".#{config[:format]}" if config[:outfile] == nil
  sdrf = SdrfXlsx.new(config)
  if config[:format] == "dot"
    open(config[:outfile],"w"){|ofh| ofh.puts sdrf.to_dot}
  else
    open("| dot -T#{config[:format]} >  #{config[:outfile]}","w"){|ofh| ofh.puts sdrf.to_dot}
  end

end
