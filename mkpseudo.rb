#!/usr/bin/ruby

# this file generates a pseudo-header from list.json.
# the resulting file is **not** a valid source file;
# it is meant to make reading the OLE classes (and their methods) easier.

require "json"
require "pp"

space = "    "
jsondata = JSON.load(File.read("list.json"))
File.open("pseudoheader.h", "w") do |fh|
  jsondata.each do |libname, klasses|
    fh << "typelib #{libname.inspect}\n{\n"
    klasses.each do |klass, khash|
      fh << space << "class #{khash["progid"].inspect}\n" << space << "{\n"
      khash["methods"].each do |name, mhash|
        if not mhash["helpstring"].nil? then
          fh << (space * 2) << "/// " << mhash["helpstring"] << "\n"
        end
        fh << (space * 2) << mhash["proto"] << ";\n\n"
      end
      fh << space << "}\n\n"
    end
    fh << "}\n\n"
  end
end
