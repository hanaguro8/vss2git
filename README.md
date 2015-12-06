vss2git.rb
==========

Outline
-------
vss2git.rb is a tool to migrate from Microsoft Visual Source Safe to Git, Mercurial or Bazaar.

vss2git.rb は、Microsoft Visual Source Safe から Git, Vercurial もしくは Bazaar に移行するためのツールです。

Reauirement
-----------
* Microsoft Windows operating system
  * Windows 7
  * Windows XP
  * Windows 8, VISTA ... not verified, may be OK

* Microsoft Visual Source Safe
  * VSS 2005 ... Language setting should be set to English.
  * VSS 6.0d ... English version is necessaly.

* Target version control system  
Command path should be added to the execution PATH.
  * Git
  * Mercurial
  * Bazaar

* Ruby 2.0 or 1.9.3 (32 bit version)

Usage
-----

    ruby vss2git.rb -r <runmode> -s <vssdir> -u <user> [-p <password>]
                    [-c <vcs>] [-d <email domain>] [-l <user list>]
                    [-b <branch>] [-e a<verbose>] [-t <time>]
                    [-w <workingdir>] VSS_PROJECT

    -r|--runmode      Run mode (0, 1, 2) (defualt:1)
                        0: Analyze
                        1: Full migration
                        2: Continuous migration
    -s|--vssdir       Absolute path to VSS repository
    -u|--user         VSS user name
    -p|--password     VSS password
    -c|--vcs          Target version control system
                      "git", "hg" or "bzr"
    -d|--emaildomain  e-mail domain
    -l|--userlist     User list file (JSON format)
                        Ex.
                        {
                          "user name on VSS":
                           ["user name on VCS", "e-mail address"],
                          "user name on VSS":
                           ["user name on VCS", "e-mail address"],
                          ...
                        }
    -b|--branch       A successful Git branching model (0, 1, 2) (default:0)
                        0: No branching model
                        1: Branching model type 1
                           master:  Production branch
                           develop: Development branch
                        2: Branching model type 2
                           master:  Development branch
                           product: Production branch
    -e|--verbose      Verbose mode (0, 1, 2) (default:1)
                        STDOUT
                          0-1: Output migration log + author list
                          2:   + dump of internal objest (for debug)
                        STDERR
                          0:   No output
                          1-2: Processing status
    -t|--timeshift    Time to shift (-12 .. 12)
    -w|--workingdir   Path to the root of working folder
    -v|--version      Print version
    -h|--help         Print help
    
Example
-------

**Case 1**

* VSS repository: *C:\vssrepo\library*
* VSS project: *$/*
* VSS user: *hanaguro*
* Target VCS: *git*

Command:
    
    > ruby vss2git.rb -r1 -s C:\vssrepo\library -u hanaguro -c git $/ >mig.log

**Case 2**

* VSS repository: *\\\vssrepo\library*
* VSS project: *$/PRODUCT-1*
* VSS user: *hanaguro*
* VSS password: *abc123*
* Target VCS: *git*
* Email domain: *mail.abcdefg.co.jp*
* User list file: *C:\doc\user.json*
* Branching model: *type 2*
* Working root: *PRODUCT-1*

Command:

    > ruby vss2git.rb -r1 -s \\\vssrepo\library -u hanaguro -p abc123
      -c git -d mail.abcdefg.co.jp -l c:\doc\user.json -b2 
      -w PRODUCT-1 $/PRODUCT-1 >mig.log

Let's try!
----------
Let's migrate from sample VSS repository to Git repository.

**1. Create sample VSS**

Enter "sample" folder and execute "mk\_sample\_vss.rb".  
The "mk\_sample\_vss.rb" do following.  

* Create working folder "sample\work".
* Create Sample VSS repository in "sample\vss".

Command:

    \> cd sample  
    sample\> ruby mk_sample_vss.rb
    Enter VSS command directory: C:\Program Files (x86)\Microsoft Visual SourceSafe

**2. Excute vss2git.rb to migrate**

Enter "sample" folder and execute "migrate.bat" to migrate.  
The "migrate.bat" do following.

* Create git working folder "sample\git".
* Migrate VSS to git.

Command:

    sample\> migrate

Change log
----------
### Ver 1.10
- Change the specification of command line option "-r".  
When you specify "-r0", you can analyze VSS to get VSS information and user list without migration.
