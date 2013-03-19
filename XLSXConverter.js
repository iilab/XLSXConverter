window.XLSXConverter = (function(_){
    
    //The propt type map is not kept in a separate JSON file because
    //origin policies might prevent it from being fetched when this script
    //is used from the local file system.
    var promptTypeMap = {
        "text" : {"type":"string"},
        "integer" : {"type":"integer"},
        "decimal" : {"type":"number"},
        "acknowledge" : {"type":"boolean"},
        "select_one" : {"type":"string"},
        "select_multiple": {
            "type": "array",
            "isPersisted": true, 
            "items" : {
                "type":"string"
            }
        },
        "select_one_with_other" : {"type":"string"},
        "geopoint" : {
            "name": "geopoint",
            "type": "object",
            "elementType": "geopoint",
            "properties": {
                "latitude": {
                    "type": "number"
                },
                "longitude": {
                    "type": "number"
                },
                "altitude": {
                    "type": "number"
                },
                "accuracy": {
                    "type": "number"
                }
            }
        },
        "barcode": {"type":"string"},
        "with_next": {"type":"string"},
        "goto": null,
        "label": null,
        "screen": null,
        "note": null,
        "error" : null,
        "image": {
            "type": "object",
            "elementType": "mimeUri",
            "isPersisted": true,
            "properties": {
                "uri": {
                    "type": "string"
                },
                "contentType": {
                    "type": "string",
                    "default": "image/*"
                }
            }
        }, 
        "audio": {
            "type": "object",
            "elementType": "mimeUri",
            "isPersisted": true,
            "properties": {
                "uri": {
                    "type": "string"
                },
                "contentType": {
                    "type": "string",
                    "default": "audio/*"
                }
            }
        }, 
        "video": {
            "type": "object",
            "elementType": "mimeUri",
            "isPersisted": true,
            "properties": {
                "uri": {
                    "type": "string"
                },
                "contentType": {
                    "type": "string",
                    "default": "video/*"
                }
            }
        },
        "date": {
            "type": "object",
            "elementType": "date"
        }, 
        "time": {
            "type": "object",
            "elementType": "time"
        }, 
        "datetime": {
            "type": "object",
            "elementType": "dateTime"
        }
    };
    var warnings = {
        __warnings__: [],
        warn: function(rowNum, message){
            //rowNum is incremented by 1 because in excel it is not 0 based
            //there might be a better place to do this.
            this.__warnings__.push("[row:"+ (rowNum + 1) +"] " + message);
        },
        clear: function(){
            this.__warnings__ = [];
        },
        toArray: function(){
            return this.__warnings__;
        }
    };
    
    var XLSXError = function(rowNum, message){
        //rowNum is incremented by 1 because in excel it is not 0 based
        //there might be a better place to do this
        return Error("[row:"+ (rowNum + 1) +"] " + message);
    };
    /*
    Extend the given object with any number of additional objects.
    If the objects have child objects with matching keys, they will
    extend eachother rather than being overwritten.
    */
    var recursiveExtend = function(obj) {
        _.each(Array.prototype.slice.call(arguments, 1), function(source) {
            if (source) {
                for (var prop in source) {
                    if (prop in obj) {
                        if (_.isObject(obj[prop]) && _.isObject(source[prop])) {
                            obj[prop] = recursiveExtend(obj[prop], source[prop]);
                        } else {
                            obj[prop] = [].concat(obj[prop]).concat(source[prop]);
                        }
                    } else {
                        obj[prop] = source[prop];
                    }
                }
            }
        });
        return obj;
    };
    /*
    [a,b,c] => {a:{b:c}}
    */
    var listToNestedDict = function(list){
        var outObj = {};
        if(list.length > 1){
            outObj[list[0]] = listToNestedDict(list.slice(1));
            return outObj;
        } else {
            return list[0];
        }
    };
    /*
    Construct a JSON object from JSON paths in the headers.
    For now only dot notation is supported.
    For example:
    {"text.english": "hello", "text.french" : "bonjour"}
    becomes
    {"text": {"english": "hello", "french" : "bonjour"}.
    */
    var groupColumnHeaders = function(row) {
        var outRow = Object.create(row.__proto__);
        _.each(row, function(value, columnHeader){
            var chComponents = columnHeader.split('.');
            outRow = recursiveExtend(outRow, listToNestedDict(chComponents.concat(value)));
        });
        return outRow;
    };
    /*
    Generates a model for ODK Survey.
    */
    var generateModel = function(prompts, promptTypeMap){
        var model = {};
        
        _.each(prompts, function(prompt){
            var schema;
            if("prompts" in prompt){
                _.extend(model, generateModel(prompt['prompts'], promptTypeMap));
            }
            if(prompt.type in promptTypeMap) {
                schema = promptTypeMap[prompt.type];
                if(schema){
                    if("name" in prompt){
                        if(prompt.name.match(" ")){
                            throw XLSXError(prompt.__rowNum__, "Prompt names can't have spaces.");
                        }
                        if(prompt.name in model){
                            warnings.warn(prompt.__rowNum__, "Duplicate name found");
                        }
                        model[prompt.name] = schema;
                    } else {
                        throw XLSXError(prompt.__rowNum__, "Missing name.");
                    }
                }
            }
        });
        return model;
    };
    
    var parsePrompts = function(sheet){
        var type_regex = /^(\w+)\s*(\S.*)?\s*$/;
        var outSheet = [];
        var outArrayStack = [{prompts : outSheet}];
        _.each(sheet, function(row){
            var curStack = outArrayStack[outArrayStack.length - 1].prompts;
            var typeParse, typeControl, typeParam;
            var outRow = row;
            //Parse the type column:
            if('type' in outRow) {
                typeParse = outRow.type.match(type_regex);
                if(typeParse && typeParse.length > 0) {
                    typeControl = typeParse[typeParse.length - 2];
                    typeParam = typeParse[typeParse.length - 1];
                    if(typeControl === "begin"){
                        outRow.prompts = [];
                        outRow.type = typeParam;
                        //Second type parse is probably not needed, it's just
                        //there incase begin ____ statements ever need a parameter
                        var secondTypeParse = outRow.type.match(type_regex);
                        if(secondTypeParse && secondTypeParse.length > 1) {
                            outRow.type = secondTypeParse[0];
                            outRow.param = secondTypeParse[1];
                        }
                        outArrayStack.push(outRow);
                    } else if(typeControl === "end"){
                        if(outArrayStack.length <= 1){
                            throw XLSXError(row.__rowNum__, "Unmatched end statement.");
                        }
                        outArrayStack.pop();
                        return;
                    } else {
                        outRow.type = typeControl;
                        outRow.param = typeParam;
                    }
                }
            }
            curStack.push(outRow);
        });
        if(outArrayStack.length > 1) {
            throw XLSXError(outArrayStack.pop().__rowNum__, "Unmatched begin statement.");
        }
        return outSheet;
    };
    
    var removeCarriageReturns = function(row){
        var outRow = Object.create(row.__proto__);
        _.each(row, function(value, key){
            if(_.isString(value)){
                outRow[key] = value.replace(/\r/g, "");
            } else {
                 outRow[key] = value;
            }
        });
        return outRow;
    };
    return {
        processJSONWorkbook : function(wbJson){
            warnings.clear();
            _.each(wbJson, function(sheet, sheetName){
                _.each(sheet, function(row, rowIdx){
                    sheet[rowIdx] = groupColumnHeaders(removeCarriageReturns(row));
                });
            });
            wbJson['survey'] = parsePrompts(wbJson['survey']);
            
            var userDefPrompts;
            if("prompt_types" in wbJson) {
                userDefPrompts = _.groupBy(wbJson["prompt_types"], "name");
                _.each(userDefPrompts, function(value, key){
                    if(_.isArray(value)){
                        userDefPrompts[key] = value[0];
                    }
                });
                wbJson["prompt_types"] = _.extend(promptTypeMap, userDefPrompts);
            }
        
            var generatedModel = generateModel(wbJson['survey'], promptTypeMap);
            var userDefModel;
            if("model" in wbJson){
                userDefModel = _.groupBy(wbJson["model"], "name");
                _.each(userDefModel, function(value, key){
                    if(_.isArray(value)){
                        userDefModel[key] = value.schema;
                    }
                });
                wbJson['model'] = _.extend(generatedModel, wbJson['model']);
            } else {
                wbJson['model'] = generatedModel;
            }
            
            if('choices' in wbJson){
                wbJson['choices'] = _.groupBy(wbJson['choices'], 'list_name');
            }
            
            return wbJson;
        },
        //Returns the warnings from the last workbook processed.
        getWarnings: function(){
            return warnings.toArray();
        }
    };
})(_);
