# = create sample VSS database
#
require 'win32ole'
require 'getoptlong'
require 'time'
require 'json'
require 'pp'

#------------------------------------------------------------------------------
# Utility
#---------------------------------------------------------------------------*/
module Utility

  # Print header line
  # "str"
  # "------------------------------------------------------------------------"
  #
  # str:: Message
  #----------------------------------------------------------------------------
  def pps_header(str)
    str = "\n" + str + "\n" + line
    pps(str)
  end

  # Dump hash object (for debug)
  # "key                     = value"
  #
  # str:: Hash name
  # obj:: Hash object
  #----------------------------------------------------------------------------
  def pps_hash(str, obj)
    pps_header(str)

    s = ""
    obj.each do |key, value|
      s += "  #{key}".ljust(24) + "= #{value}\n"
    end
    pps(s)
  end

  # Dump object (for debug)
  #
  # str:: Object name
  # obj:: Object
  #----------------------------------------------------------------------------
  def pps_object(str, obj)
    pps_header(str)
    pp obj
  end

  # Print message to STDOUT
  #
  # str:: Message
  #----------------------------------------------------------------------------
  def pps(str)
    puts "#{str}"
  end
  private :pps

  # Seperlator line
  # "------------------------------------------------------------------------"
  #----------------------------------------------------------------------------
  def line
    "-" * 72
  end
  private :line

  # Print processing status to STDERR when verbose != 0
  # "mes1                            : mes2"
  #
  # verbose:: verbose mode
  # mes1::    1st message
  # mes2::    2nd message
  #----------------------------------------------------------------------------
  def ppe_status(verbose, mes1, mes2 = "")
    return if verbose == 0

    s =  mes1.ljust(32)
    s += ": " + mes2 unless mes2 == ""
    s += "\r"
    STDERR.print s
  end

  # Print error message and exit
  #
  # str:: Error message
  #----------------------------------------------------------------------------
  def ppe_exit(str)
    STDERR.print(str)
    exit 1
  end

  # Execute windows command
  #
  # str:: Command string
  #----------------------------------------------------------------------------
  def ex(str)
    puts str
    %x(#{str})
  end
end


#------------------------------------------------------------------------------
# CLASS: Vss
#
# VSS (Visual Source Safe) operation class
#
# This class operates Visual Source Safe via Win32OLE interface.
#------------------------------------------------------------------------------
class VssError < StandardError; end

class Vss
  include Utility

  # VSS constant
  module VssConstant
  end

  #-- definition of VCS oparations
  ACTION_ADD    = "ADD"
  ACTION_BRANCH = "BRANCH"
  ACTION_COMMIT = "COMMIT"
  ACTION_INIT   = "INIT"
  ACTION_MERGE  = "MERGE"
  ACTION_MOVE   = "MOVE"
  ACTION_REMOVE = "REMOVE"
  ACTION_TAG    = "TAG"

  # VSS actions
  VSS_ACTION = [
    { VssAction: /added/,                Sym: :Added,              Action: nil        },
    { VssAction: /archived versions of/, Sym: :ArchivedVersionsOf, Action: ACTION_ADD },
    { VssAction: /archived/,             Sym: :Archived,           Action: nil        },
    { VssAction: /branched at version/,  Sym: :BranchedAtVersion,  Action: nil        },
    { VssAction: /checked in/,           Sym: :CheckedIn,          Action: ACTION_ADD },
    { VssAction: /created/,              Sym: :Created,            Action: ACTION_ADD },
    { VssAction: /deleted/,              Sym: :Deleted,            Action: nil        },
    { VssAction: /destroyed/,            Sym: :Destroyed,          Action: nil        },
    { VssAction: /labeled/,              Sym: :Labeled,            Action: ACTION_TAG },
    { VssAction: /moved from/,           Sym: :MovedFrom,          Action: nil        },
    { VssAction: /moved to/,             Sym: :MovedTo,            Action: nil        },
    { VssAction: /pinned to version/,    Sym: :PinnedToVersion,    Action: nil        },
    { VssAction: /purged/,               Sym: :Purged,             Action: nil        },
    { VssAction: /recovered/,            Sym: :Recovered,          Action: nil        },
    { VssAction: /renamed to/,           Sym: :RenamedTo,          Action: nil        },
    { VssAction: /restored/,             Sym: :Restored,           Action: nil        },
    { VssAction: /rollback to version/,  Sym: :RollbackToVersion,  Action: nil        },
    { VssAction: /shared/,               Sym: :Shared,             Action: nil        },
    { VssAction: /unpinned/,             Sym: :Unpinned,           Action: nil        },
    { VssAction: /.*/,                   Sym: :Other,              Action: "OTHER"    }]

  # Initialize instance
  #
  # vssdir::   VSS database directory
  # user::     VSS user name
  # password:: VSS pasword
  # project::  Project name
  # workdir::  Working directory
  # verbose::  Verbose mode
  #--------------------------------------------------------------------------
  def initialize(vssdir, user, password, project, workingdir, verbose)
    @vssdir     = vssdir
    @project    = project
    @workingdir = workingdir
    @verbose    = verbose

    # open VSS db
    begin
      @vssdb = WIN32OLE.new("SourceSafe")
      WIN32OLE.const_load(@vssdb, VssConstant)
    rescue
      raise VssError,
        %(\nERROR: Visual Source Safe is not installed.)
    end

    # validation of srcsafe.ini
    file = vssdir + "srcsafe.ini"
    unless FileTest.exist?(file)
      raise VssError,
        %(\nERROR: Invalid VSS database folder: #{file})
    end

    begin
      @vssdb.Open(file, user, password)
    rescue
      raise VssError,
        %(\nERROR: Invalid user name: #{user})
    end

    if @vssdb.GetSetting("Force_Dir") == "Yes"
      raise VssError,
        %(\nERROR: VSS setting error.) +
        %(\n  "Assume working folder based on current project" should be off.) +
        %(\n  "Tools" -> "Options" -> "Command Line Options" tab)
    end

    if @vssdb.GetSetting("Force_Prj") == "Yes"
      raise VssError,
        %(\nERROR: VSS setting error.) +
        %(\n  "Assume project based on working folder" should be off.) +
        %(\n  "Tools" -> "Options" -> "Command Line Options" tab)
    end
  end

  # Get history
  #----------------------------------------------------------------------------
  def get_history
    history  = []
    @counter = { File: 0 }
    @vssinfo = {}
    VSS_ACTION.each do |act|
      @vssinfo[act[:Sym]] = 0
    end

    files = get_filelist(@project)

    files.each_with_index do |file, i|
      item = get_item(file)
      next unless item

      history += get_history_of_the_file(file, item)
      ppe_status(
        @verbose,
        "Get history...",
        "#{i + 1} / #{files.size} files")
    end

    ppe_status(@verbose, "\n")
    pps_object("history generated by get_history", history) if @verbose >= 3

    [history, @vssinfo]
  end

  # Get file list
  #
  # project:: project folder name (Ex. $/, $/project etc.)
  #----------------------------------------------------------------------------
  def get_filelist(project)
    files = walk_tree(project).uniq.sort
    ppe_status(@verbose, "Make file list...", "#{files.size} files\n")
    files
  end
  private :get_filelist

  # Walk VSS project tree to get file list
  #
  # project:: project folder name (Ex. $/, $/project etc.)
  #----------------------------------------------------------------------------
  def walk_tree(project)
    files = []

    root = get_item(project)
    return files unless root

    items = root.Items(false)

    items.each do |item|
      pps_item(item)

      if item.Type == VssConstant::VSSITEM_PROJECT
        subproject = item.Name
        files += walk_tree("#{project}#{subproject}/")
        ppe_status(
          @verbose, "Meke file list...", "#{@counter[:File]} files")
      else
        @counter[:File] += 1
        files << item.Spec
      end
    end

    files
  end
  private :walk_tree

  # Get IVSSItem object
  #
  # file:: file name or project folder name
  #----------------------------------------------------------------------------
  def get_item(file)
    begin
      @vssdb.VSSItem(file, false)
    rescue
      message = "WARNING: Cannot handle the file: #{file}"
      puts message
      ppe_status(@verbose, message + "\n")
      nil
    end
  end
  private :get_item

  # Get history of the file
  #
  # file:: file name
  # item:: VSSItem object
  #----------------------------------------------------------------------------
  def get_history_of_the_file(file, item)
    history = []

    versions = item.Versions

    versions.each do |ver|
      pps_version(file, ver)

      hs = {}
      hs[:File]          = file
      hs[:Version]       = ver.VersionNumber
      hs[:Author]        = ver.Username.downcase
      hs[:Date]          = ver.Date
      hs[:Message]       = ver.Comment.gsub(/[\r\n][\r\n]/, "\n").chomp
      hs[:Tag]           = ver.Label.chomp
      hs[:LatestVersion] = (item.VersionNumber == hs[:Version])

      pps ver.Action if @verbose >= 3

      action = VSS_ACTION.find do |act|
        ver.Action.downcase =~ act[:VssAction]
      end

      @vssinfo[action[:Sym]] += 1

      case action[:Action]
      when "OTHER"
        raise VssError,
          %(\nERROR: Unexpected operation: #{ver.Action.downcase})
      when nil
        next
      else
        hs[:Action] = action[:Action]
      end

      history << hs
    end
    history
  end
  private :get_history_of_the_file

  # Get file
  #
  # file::    File name or project folder name
  # version:: Version number.
  #           If version is nil, the latest version of file is got.
  #----------------------------------------------------------------------------
  def get_file(file, version)
    item = get_item(file)
    return unless item

    item = item.Version(version) if version
    fail "item.Version() return nil" unless item

    # VSSFLAG setting
    flag = VssConstant::VSSFLAG_CMPFAIL |
      VssConstant::VSSFLAG_FORCEDIRNO |
      VssConstant::VSSFLAG_RECURSYES |
      VssConstant::VSSFLAG_REPREPLACE

    # build local path name
    lpath = file.gsub(/[^\/]*$/, "").gsub(@project, "")
    lpath = @workingdir + lpath.gsub(/\//, "\\")
    lpath.gsub!(/\\$/, "")

    case item.Type
    when VssConstant::VSSITEM_PROJECT
      lpath = ".\\" + lpath
      ex("if not exist #{lpath} md #{lpath}")
    when VssConstant::VSSITEM_FILE
      lpath = ".\\" + lpath + "\\" + item.Name
    end

    # get file
    puts "Checkout #{file}"
    begin
      item.Get(lpath, flag)
      true
    rescue
      puts "WARNING: Cannot get file: #{file}: V#{version}"
      false
    end
  end

  # Dump IVSSItem object (for debug)
  #
  # item:: IVSSItem object
  #----------------------------------------------------------------------------
  def pps_item(item)
    return unless @verbose >= 3

    pps_header("item")
    puts "item.Spec: #{item.Spec}"
    puts "item.Deleted: #{item.Deleted}"
    puts "item.Type: #{item.Type}"
    puts "item.LocalSpec: #{item.LocalSpec}"
    puts "item.Name: #{item.Name}"
    puts "item.VersionNumber: #{item.VersionNumber}"
  end
  private :pps_item

  # Dump IVSSVersion object (for debug)
  #
  # file:: File name
  # ver::  IVSSVersion object
  #----------------------------------------------------------------------------
  def pps_version(file, ver)
    return unless @verbose >= 3

    pps_header("version")
    puts "file: #{file}"
    puts "ver.VersionNumber: #{ver.VersionNumber}"
    puts "ver.Action: #{ver.Action}"
    puts "ver.Date: #{ver.Date}"
    puts "ver.Username: #{ver.Username}"
    puts "ver.Comment: #{ver.Comment}"
    puts "ver.Label: #{ver.Label}"
  end
  private :pps_version

  # Make sample VSS database
  #-----------------------------------------------------------------------------
  def make_sample_vss
    # VSSFLAG setting
    flag = VssConstant::VSSFLAG_CMPFAIL |
      VssConstant::VSSFLAG_FORCEDIRNO |
      VssConstant::VSSFLAG_RECURSYES |
      VssConstant::VSSFLAG_REPREPLACE |
      VssConstant::VSSFLAG_GETYES

    cnt = 4

    # add VSS user
    puts "Add VSS user account..."
    (0...cnt).each do |i|
      begin
        user = "user#{i}.user"
        @vssdb.AddUser(user, "", false)
      rescue => e
        print e.message
      end
    end

    # create initial file
    puts "Create initial file..."
    (0...cnt).each do |i|
      (0...cnt).each do |j|
        dir = "dir#{i}\\dir#{j}"
        %x(if not exist #{dir} md #{dir})

        (0...cnt).each do |k|
          file = "#{dir}\\file#{k}.txt"

          %x(echo off&&if not exist #{file} echo #{k}:>#{file})
        end
      end
    end

    # initial commit
    puts "Initial check-in..."
    vssItem = @vssdb.VSSItem("$/", false)
    begin
      vssItem.Add(".", "Initial commit", flag)
    rescue => e
      print e.message
    end

    db = WIN32OLE.new("SourceSafe")
    vssdir = @vssdir + "scrsafe.ini"

    # check-out, modify and check-in
    (0...cnt).each do |i|

      puts "Check-out, modify and check-in..."
      (0...cnt).each do |j|
        user = "user#{j}.user"
        dir = "dir#{i}\\dir#{j}"

        db.Open(vssdir, user, "")
        vssItem = db.VSSItem("$/dir#{i}/dir#{j}", false)

        vssItem.Checkout("", dir, flag)

        (0...cnt).each do |k|
          file = "#{dir}\\file#{k}.txt"

          %x(echo off&&attrib -R #{file}&&echo Modified by #{user}>>#{file})
        end

        vssItem.Checkin("Commit by #{user}", dir, flag)
        db.Close
      end

      sleep(2)
      tag = "Ver#{i}.0"
      puts "Tag: #{tag}"
      db.Open(vssdir, "admin", "")
      vssItem = db.VSSItem("$/", false)
      vssItem.Label(tag)
      db.Close
    end
  end
end # Class Vss




#j enter VSS command directory
print "Enter VSS command directory: "
vsspath = ARGF.gets.chop

mkss = vsspath + "\\" + "mkss.exe"
unless File.exist?(mkss)
  puts "VSS command (mkss.exe) is not found."
  puts "(#{mkss})"
  exit 1
end

#j create working directory
puts "Create working directory..."
%x(if exist vss rd vss /s /q)
%x(md vss)
%x(if exist work rd work /s /q)
%x(md work)

#j create VSS database
puts "Create VSS database..."
%x(\"#{mkss}"\ vss /V6)

#j crate sample source file and check-in VSS
saveddir = Dir.pwd
Dir.chdir "work"
begin
  vss = Vss.new("..\\vss\\", "admin", "", "$/", "", 1)
  vss.make_sample_vss
ensure
  Dir.chdir saveddir
end

puts "Sample VSS database is created in vss directory."
