#!/usr/bin/ruby --disable-gems

require "json"
require "win32ole"

### -- options begin here ---
# set to true to avoid emitting referenced classes (i.e., duplicates)
AVOID_DUPLICATE_CLASSES = true
# whether to shorten the param array (from a hash to an array being [<param_type>, <param_name>, <default_value>])
# if there is no default value, default_value will be nil.
SIMPLIFIED_PARAM_ARRAY = true
### --- options end here ---

jsondata = {}
guids = {}
$stdout.sync = true

begin
  WIN32OLE_TYPE.typelibs.each { |typelib|
    if typelib.length > 0 then
      begin
        klasses = WIN32OLE_TYPE.ole_classes(typelib)
        if klasses.length > 0 then
          puts "processing #{typelib.inspect} ..."
          typhash = {"name" => typelib, "classes" => {}, "path" => klasses[0].ole_typelib.path}
          klasses.each { |klass|
            klasshash = {"methods" => {}}
            if klass.visible? then
              methods = klass.ole_methods
              if (AVOID_DUPLICATE_CLASSES && (not guids.key?(klass.guid))) then
                if (klass.progid != nil) && (methods.length > 0) then
                  puts "  processing #{klass.name} (#{klass.progid.inspect}) ..."
                  guids[klass.guid] = true
                  klasshash["name"] = klass.name
                  klasshash["progid"] = klass.progid
                  klasshash["guid"] = klass.guid
                  klasshash["version"] = "#{klass.major_version}.#{klass.minor_version}"
                  if klass.helpstring.length > 0 then
                    klasshash["helpstring"] = klass.helpstring
                  end
                  methods.each { |meth|
                    methhash = {"params" => []}
                    name = meth.name
                    if meth.visible? then
                      invkind = meth.invoke_kind.downcase
                      rt = meth.return_type
                      helpstr = meth.helpstring.strip
                      helpfile = meth.helpfile.strip
                      #puts "    processing #{name.inspect} ..."
                      methhash["return_type"] = rt
                      methhash["invkind"] = invkind
                      # build prototype params string, and populate "params" array
                      protoparams = meth.params.map { |arg|
                        ret = String.new
                        if SIMPLIFIED_PARAM_ARRAY then
                          parm = [arg.ole_type, arg.name, nil]
                        else
                          parm = {}
                          parm["name"] = arg.name
                          parm["type"] = arg.ole_type
                          parm["is_input"] = arg.input?
                          parm["is_output"] = arg.output?
                          parm["is_optional"] = arg.optional?
                        end
                        iospec = if arg.input? then "_in_" elsif arg.output? then "_out_" end
                        iospec = if iospec == nil then "" else (iospec + " ") end
                        ret << arg.ole_type << " "  << iospec << arg.name
                        # needed because #<< treats integers oddly, for some reason
                        if arg.default != nil then
                          str = arg.default.to_s
                          if str.length > 0 then
                            if SIMPLIFIED_PARAM_ARRAY then
                              parm[3] = arg.default
                            else
                              parm["default"] = arg.default
                            end
                            ret << "=" << str
                          end
                        end
                        methhash["params"] << parm
                        next ret
                      }.join(", ")
                      if methhash["params"].empty? then
                        methhash["params"] = nil
                      end
                      if helpstr.length > 0 then
                        methhash["helpstring"] = helpstr
                      end
                      # build function prototype
                      begin
                        methhash["proto"] = case invkind
                          when "func" then
                            "#{rt} #{name}(#{protoparams})"
                          when "propertyget" then
                            "property #{rt} #{name} {get;}"
                          when "propertyput", "propertyputref" then
                            "property #{rt} #{name} {get;set;}"
                          else
                            raise Exception, "unknown invkind #{invkind.inspect}"
                        end
                      end
                      klasshash["methods"][name] = methhash
                    end
                  }
                end
                if not klasshash["methods"].empty? then
                  klasshash["methods"] = klasshash["methods"].sort_by{|name, _| name}.to_h
                  typhash["classes"][klass.name] = klasshash
                end
              end
            end
          }
          if not typhash.empty? then
            if not typhash["classes"].empty? then
              jsondata[typelib] = typhash["classes"].sort_by{|classname, _| classname}.to_h
            end
          end
        end
      rescue WIN32OLERuntimeError => err
        #$stderr << "error loading #{typelib.inspect}: " <<  err << "\n"
      end
    end
  }
ensure
  jsondata = jsondata.sort_by{|libname, _| libname}.to_h
  File.open("list.json", "w") do |fh|
    fh << JSON.pretty_generate(jsondata)
    fh << "\n"
  end
end
