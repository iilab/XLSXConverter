(function() {

    // Establish the root object, `window` in the browser, or `global` on the server.
    var root = this;

    var _ = root._;

    //The propt type map is not kept in a separate JSON file because
    //origin policies might prevent it from being fetched when this script
    //is used from the local file system.
    var promptTypeMap = {


        // yes_no,yes_no_unknown,single_line_text,
        // multi_line_text,number,check_boxes,radio_buttons,drop_down,date,
        "single_line_text": {
            "type": "string"
        },
        "multi_line_text": {
            "type": "string"
        },
        "number": {
            "type": "string"
        },
        "yes_no": {
            "type": "string",
        },
        "yes_no_unknown": {
            "type": "string",
        },
        "drop_down": {
            "type": "select"
        },
        "check_boxes": {
            "type": "string"
        },
        "radio_buttons": {
            "type": "string"
        },

        "date": {
            "type": "string",
        }
    };
    var warnings = {
        __warnings__: [],
        warn: function(rowNum, message) {
            //rowNum is incremented by 1 because in excel it is not 0 based
            //there might be a better place to do this.
            this.__warnings__.push("[row:" + (rowNum + 1) + "] " + message);
        },
        clear: function() {
            this.__warnings__ = [];
        },
        toArray: function() {
            return this.__warnings__;
        }
    };

    var XLSXError = function(rowNum, message) {
        //rowNum is incremented by 1 because in excel it is not 0 based
        //there might be a better place to do this
        return Error("[row:" + (rowNum + 1) + "] " + message);
    };
    /*
    Extend the given object with any number of additional objects.
    If the objects have matching keys, the values of those keys will be
    recursively merged, either by extending eachother if any are objects,
    or by being combined into an array if none are objects.
    */
    var recursiveExtend = function(obj) {
        _.each(Array.prototype.slice.call(arguments, 1), function(source) {
            if (source) {
                for (var prop in source) {
                    if (prop in obj) {
                        if (_.isObject(obj[prop]) || _.isObject(source[prop])) {
                            //If one of the values is not an object,
                            //put it in an object under the key "default"
                            //so it can be merged.
                            if (!_.isObject(obj[prop])) {
                                obj[prop] = {
                                    "default": obj[prop]
                                };
                            }
                            if (!_.isObject(source[prop])) {
                                source[prop] = {
                                    "default": source[prop]
                                };
                            }
                            obj[prop] = recursiveExtend(obj[prop], source[prop]);
                        } else {
                            //If neither value is an object put them in an array.
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
    var listToNestedDict = function(list) {
        var outObj = {};
        if (list.length > 1) {
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
        var outRow = Object.create(row.__proto__ || row.prototype);
        _.each(row, function(value, columnHeader) {
            var chComponents = columnHeader.split('.');
            outRow = recursiveExtend(outRow, listToNestedDict(chComponents.concat(value)));
        });
        return outRow;

    };
    /*
    Generates a model for Alpaca.
    */
    var generateAlpaca = function(formList, promptTypeMap) {
        var models = [];
        _.each(formList, function(form){
            //create a schema and options object for this form
            var schema = {};
            var options = {};

            var beginRow = form.shift();
            //make sure that there is a begin_form marker at the top
            if(beginRow.type == "begin_form"){
                //establish the schema basic items
                schema.title = beginRow.en_labels;
                schema._id = beginRow.id;
                schema.type = "object";
                schema.properties = {};

                //establish options basic items
                options.fields = {};

                if(beginRow.en_help){
                    schema.description = beginRow.en_help;
                }
                for(var i=0;i<form.length;i++){
                    //grab single form item in form
                    var formItem = form[i];
                    var itemType = promptTypeMap[formItem.type];
                    //create a schema and options object to match it
                    var schemaObj = {};
                    var optionsObj = {};
                    //make sure the formItem has an id and legit type
                    if(formItem.id != undefined && itemType != undefined){

                        optionsObj.label = formItem.en_labels
                        optionsObj.helper = formItem.sw_labels;
                        schemaObj.type = "string";
                        if(formItem.constraint != undefined){
                            switch(formItem.constraint){
                                case "required":
                                    schemaObj.required = true;
                                    break;
                            }
                        }
                        if(itemType.type != "string"){
                            // optionsObj.type = itemType.type;
                        }

                        //add form items to schema and options
                        schema.properties[formItem.id] = schemaObj;
                        options.fields[formItem.id] = optionsObj

                    }
                }
            }
            var model = {"schema": schema, "options": options};
            console.log(JSON.stringify(model));
            models.push(model);
        });
        return models;
    };

    // Cut the xlsx file into separate forms.
    var parseForms = function(sheet) {
        var type_regex = /^(\w+)\s*(\S.*)?\s*$/;
        var outSheet = [];
        _.each(sheet, function(row){
            var currStackIndex = outSheet.length-1;
            var typeMatch, typeControl;
            //parse the type
            if('type' in row){
                var outRow = row;
                typeMatch = row.type.match(type_regex);
                if(typeMatch && typeMatch.length > 0){
                    typeControl = typeMatch[0];
                    if(typeControl === "begin_form"){
                        outSheet.push([outRow]);
                    }else if(typeControl === "end_form"){
                        if(outSheet.length < 1){
                            throw XLSXError(row.__rowNum__, "Unmatched end statement.");
                        }
                    }else{
                        if(currStackIndex > -1){
                            outRow.type = typeControl;
                            outSheet[currStackIndex].push(outRow);   
                        }
                    }
                }
            }
        });
        console.log(outSheet);
        return outSheet;
    };

    //Remove carriage returns, trim values.
    var cleanValues = function(row) {
        var outRow = Object.create(row.__proto__ || row.prototype);
        _.each(row, function(value, key) {
            if (_.isString(value)) {
                value = value.replace(/\r/g, "");
                value = value.trim();
            }
            outRow[key] = value;
        });
        return outRow;
    };

    root.XLSXConverter = {
        processJSONWorkbook: function(wbJson) {
            warnings.clear();
            _.each(wbJson, function(sheet, sheetName) {
                _.each(sheet, function(row, rowIdx) {
                    var reRow = groupColumnHeaders(cleanValues(row));
                    reRow._rowNum = reRow.__rowNum__ + 1;
                    sheet[rowIdx] = reRow;
                });
            });

            //Process sheet names by converting from json paths to nested objects.
            //Sheet names become objects containing the rows in the sheet.
            var tempWb = {};
            _.each(wbJson, function(sheet, sheetName) {
                var tokens = sheetName.split('.');
                var sheetObj = {};
                sheetObj[tokens[0]] = listToNestedDict(tokens.slice(1).concat([sheet]));
                recursiveExtend(tempWb, sheetObj);
            });
            wbJson = tempWb;

            if (!('survey' in wbJson)) {
                throw Error("Missing survey sheet");
            }

            if (_.isObject(wbJson['survey'])) {
                //If the survey sheet is an object rather than an array,
                //We have multiple sheets of the form survey.x survey.y ... 
                //So we concatenate them into an alphabetically sorted array.
                wbJson['survey'] = _.flatten(_.sortBy(wbJson['survey'],
                    function(val, key) {
                        return key;
                    }), true);
            }

            wbJson['survey'] = parseForms(wbJson['survey']);

            if ('choices' in wbJson) {
                // lists is the sheet name. list_id is the column name on that sheet
                wbJson['lists'] = _.groupBy(wbJson['lists'], 'list_id');
            }

            //Generate a model:
            var userDefPrompts = {};
            // if ("prompt_types" in wbJson) {
            //     userDefPrompts = _.groupBy(wbJson["prompt_types"], "name");
            //     _.each(userDefPrompts, function(value, key) {
            //         if (_.isArray(value)) {
            //             userDefPrompts[key] = value[0].schema;
            //         }
            //     });
            // }
            var extendedPTM = _.extend({}, promptTypeMap, userDefPrompts);

            // Converts the 'survey' sheet into custom format 
            var generatedModel = generateAlpaca(wbJson['survey'], extendedPTM);
            // var userDefModel;
            // if ("model" in wbJson) {
            //     userDefModel = _.groupBy(wbJson["model"], "name");
            //     _.each(userDefModel, function(value, key) {
            //         if (_.isArray(value)) {
            //             userDefModel[key] = value[0].schema;
            //         }
            //     });
            //     wbJson['model'] = _.extend(generatedModel, userDefModel);
            // } else {
            wbJson['model'] = generatedModel;
            // }

            return wbJson['model'];
        },
        //Returns the warnings from the last workbook processed.
        getWarnings: function() {
            return warnings.toArray();
        }
    };
}).call(this);