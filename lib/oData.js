"use strict";

require("simple-errors");
var url             = require("url");
var xpath           = require("xpath");
var domParser       = new (require("xmldom").DOMParser)();
var constants       = require("constants");
var request         = require("request");
var Agentkeepalive  = require("agentkeepalive");
var _               = require("lodash");

// this class implements all features
module.exports.createClient = function (settings) {

    var whitelistProps = ["url", "timeout", "strictSSL", "streaming"];

    // Arguments validation
    if (!settings || typeof settings !== "object") throw Error.create("'settings' argument must be an object instance.");
    if (!settings.url || typeof settings.url !== "string") throw Error.create("'settings.url' property is a required string.");

    if (settings.timeout !== undefined && typeof settings.timeout !== "number") throw Error.create.create("'settings.timeout' property must be a number.");
    if (settings.strictSSL !== undefined && typeof settings.strictSSL !== "boolean") throw Error.create("'settings.strictSSL' property must be a boolean.");
    if (settings.streaming !== undefined && typeof settings.streaming !== "boolean") throw Error.create("'settings.streaming' property must be a boolean.");

    var config = _.pick(settings, whitelistProps);          // filter props in whitelist
    config.logger = settings.logger || require("winston");  // default winston
    config.security = settings.security || null;            // default no security
    config.timeout = config.timeout || 180000;              // default to 3 minutes
    config.useHTTPS = url.parse(settings.url).protocol.toUpperCase() === "HTTPS";
    return new OData(config);
};

function defaultCb(err) {
    if (err) throw err;
};

function buildPath(href, options) {
    if (options.location) return options.location;
    var command = (options.command || "");
    while (command.indexOf("//") > -1) {
        command = command.replace("//", "/");
    }
    //make sure only one "/" between href and command.
    command = command[0] === "/" ? command : "/" + command;
    href = href.replace(/\/$/, '');
    var path = href + command;
    return encodeURI(path);
};
module.exports.buildPath = buildPath;

function OData(config) {
    var self            = this;     // Auto reference
    var entitySets      = null;     // String array containing all entity sets names
    var pendingHook     = null;     // Function that will be executed after 'entitySets' array was populated.

    // sends oData requests to the server
    function requestOData(options, cb) {
        options.headers = options.headers || {};

        var dataType = (options.headers["content-type"] && (options.headers["content-type"].indexOf("xml") > -1)) ? "XML" : "JSON";

        var defaultOptions = {
            headers: {
                MaxDataServiceVersion: "3.0",
                accept: "application/json;odata=light;q=1,application/json;odata=verbose;q=0.5"
            },
            method: "GET",
            url: buildPath(config.url, options),
            data: undefined,
            json: dataType === "JSON",
            agent: config.useHTTPS ? new Agentkeepalive.HttpsAgent() : new Agentkeepalive(),
            timeout: config.timeout
        };
        var reqOptions = _.extend(defaultOptions, options);
        if (config.useHTTPS) {
            reqOptions.secureOptions = constants.SSL_OP_NO_TLSv1_2;
            reqOptions.ciphers = "ECDHE-RSA-AES256-SHA:AES256-SHA:RC4-SHA:RC4:HIGH:!MD5:!aNULL:!EDH:!AESGCM";
            reqOptions.honorCipherOrder = true;
            reqOptions.strictSSL = config.strictSSL;
        }

        if (options.data) reqOptions.headers["content-type"] = options.headers["content-type"] || "application/json;odata=verbose";

        if (config.security) config.security.addOptions(reqOptions);

        var returnStream = (options.streaming === false || options.streaming === true) ? options.streaming : !!(config.streaming);
        if (returnStream) return cb(null, request(reqOptions));


        request(reqOptions, function (err, res, body) {
            if (err) return cb(err);
            parseResponse(res, body, cb);
        });
    }


    function parseResponse (res, body, cb) {
        var result = { statusCode: res.statusCode };
        var type = res.headers["Content-Type"] || res.headers["content-type"];

        try {
            if (type && type.indexOf("xml") > -1) {
                // returns parsed XML
                result.data = domParser.parseFromString(body);
            } else if (type && type.indexOf("json") > -1) {
                // returns parsed JSON
                if (typeof body === 'object') {
                    result.data = body;
                } else {
                    result.data = JSON.parse(body);
                }
            } else {
                // returns plain body
                result.data = body;
            }
        } catch (e) {
            config.logger.debug("Unable to parse response, statusCode: ", res.statusCode);
            config.logger.debug("Unable to parse response, body: ", body);
            config.logger.debug("Unable to parse response, headers: ", res.headers);
            return cb(Error.create("Couldn't parse the response from the server. ", e));
        }
        cb(null, result);
    }
    // serialize the entity id
    function buildODataId(id) {
        if (typeof id === "string") return "'" + id + "'";
        return String(id);
    }

    // retrives the entity sets from the service's metadata
    function getEntitySets(options, cb) {
        // retrieves entity sets only once
        if (entitySets) { return cb(null, entitySets); }

        var odataOptions = _.cloneDeep(options);
        odataOptions.command = "$metadata";
        odataOptions.headers = {
            accept: "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        };

        // gets the metadata XML
        self.oData(odataOptions, function (err, result) {
            try {
                if (err) {
                    // If server does not support $metadata, the entityset will be setted to an empty string
                    entitySets = [];
                } else {
                    entitySets = xpath
                        .select("//*[local-name(.)='EntitySet']/@Name", result.data)
                        .map(function (attr) { return attr.value; });
                }
            } catch (e) {
                // process error
                entitySets = null;
                return cb(Error.create("Couldn't parse the entity sets from the metadata. ", e));
            }
            // adds methods dinamically
            if (pendingHook) {
                try {
                    pendingHook();
                } catch (e) {
                    return cb(Error.create("Couldn't add operations from metadata. ", e));
                }
            }
            // returns entities
            cb(null, entitySets);
        });
    }

    // Executes a oData Command
    this.oData = function (options, cb) {
        // handles optional 'options' argument
        if (!cb && typeof options === "function") {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};
        if (typeof options !== "object") { return cb(new Error("'options' argument is missing or invalid.")); }
        if (!options.command || typeof options.command !== "string") { return cb(new Error("options.command' argument is missing or invalid.")); }

        requestOData(options, cb);
    };

    // Returns an string array with the name of all entity sets
    this.entitySets = function (options, cb) {
        // handles optional 'options' argument
        if (!cb && typeof options === "function") {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};
        if (!typeof options !== "object") { return cb(new Error("'options' argument is missing or invalid.")); }

        getEntitySets(options, cb);
    };

    // gets an entity by its id
    this.get = function (options, cb) {
        // handles optional 'options' argument
        if (!cb && typeof options === "function") {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        if (!options.resource || typeof options.resource !== "string") { return cb(new Error("'resource' property is missing or invalid.")); }
        if (options.id === null || options.id === undefined) { return cb(new Error("'id' property is missing.")); }

        self.oData({
            method: "GET",
            command: "/" + options.resource + "(" + buildODataId(options.id) + ")"
        }, cb);
    };

    // executes a query on an entity set
    this.query = function (options, cb) {
        // handles optional 'options' argument
        if (!cb && typeof options === "function") {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};
        if (!options.resource || typeof options.resource !== "string") { return cb(new Error("'resource' property is missing or invalid.")); }

        var err = null;
        var params = ["filter", "expand", "select", "orderBy", "top", "skip"]
        .map(function (prop) {
            if (options[prop]) {
                var value = options[prop];
                if (typeof value !== "string") {
                    err = new Error("The property '" + prop + "' must be a valid string.");
                    return null;
                }
                return "$" + prop.toLowerCase() + "=" + value;
            }
            return null;
        })
        .filter(function (param) { return !!param; });

        if (err) { return cb(err); }

        params.push("$inlinecount=" + (options.inLineCount ? "allpages" : "none"));

        self.oData({
            method: "GET",
            command: "/" + options.resource + (params.length === 0 ? "" : "?" + params.join("&"))
        }, cb);
    };

    this.download = function (options, cb) {
        // handles optional 'options' argument
        if (!cb && typeof options === "function") {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};
        if (!options.resource || typeof options.resource !== "string") { return cb(new Error("'resource' property is missing or invalid.")); }

        self.oData({
            method: "GET",
            command: "/" + options.resource + "/File"
        }, function (err, data) {
            if (err) { return cb(err); }

            var file = data && data.data && data.data.d && data.data.d;
            if (!file || !file.ServerRelativeUrl) return (new Error("File location not found."));

            var reqOptions = {
                method: "GET",
                streaming: true,
                command: file.ServerRelativeUrl,
                location: file.ServerRelativeUrl
            };
            self.oData(reqOptions, function (errStream, stream) {
                if (errStream) { return errStream; }
                cb(null, stream, { "Content-Disposition": "attachment;filename="   + file.Name });
            });
        });
    };

    this.processQuery = function (options, cb) {
        // handles optional 'options' argument
        if (!cb && typeof options === "function") {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};

        var odataOptions = {
            method: "POST",
            data: options.data,
            command: "/ProcessQuery",
            headers: {
                "content-type": "text/xml"
            }
        };
        self.oData(odataOptions, cb);
    };

    // gets all links of an entity instance to entities of a specific type
    this.links = function (options, cb) {
        // handles optional 'options' argument
        if (!cb && typeof options === "function") {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};
        if (options.id === null || options.id === undefined) { return cb(new Error("'id' property is missing.")); }
        if (!options.resource || typeof options.resource !== "string") { return cb(new Error("'resource' property is missing or invalid.")); }
        if (!options.entity || typeof options.entity !== "string") { return cb(new Error("'entity' property is missing or invalid.")); }

        self.oData({
            method: "GET",
            command: "/" + options.resource + "(" + buildODataId(options.id) + ")/$links/" + options.entity
        }, cb);
    };

    // returns the number of elements of an entity set
    this.count = function (options, cb) {
        // handles optional 'options' argument
        if (!cb && typeof options === "function") {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};
        if (!options.resource || typeof options.resource !== "string") { return cb(new Error("'resource' property is missing or invalid.")); }

        self.oData({
            method: "GET",
            command: "/" + options.resource + "/$count"
        }, cb);
    };

    // adds an entity instance to an entity set
    this.create = function (options, cb) {
        // handles optional 'options' argument
        if (!cb && typeof options === "function") {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};
        if (options.data === null || options.data === undefined) { return cb(new Error("'data' property is missing.")); }
        if (!options.resource || typeof options.resource !== "string") { return cb(new Error("'resource' property is missing or invalid.")); }

        self.oData({
            data: options.data,
            method: "POST",
            command: "/" + options.resource
        }, cb);
    };

    // does a partial update of an existing entity instance
    this.replace = function (options, cb) {
        // handles optional 'options' argument
        if (!cb && typeof options === "function") {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};
        if (options.id === null || options.id === undefined) { return cb(new Error("'id' property is missing.")); }
        if (options.data === null || options.data === undefined) { return cb(new Error("'data' property is missing.")); }
        if (!options.resource || typeof options.resource !== "string") { return cb(new Error("'resource' property is missing or invalid.")); }

        self.oData({
            data: options.data,
            method: "PUT",
            command: "/" + options.resource + "(" + buildODataId(options.id) + ")",
            etag: "*"
        }, cb);
    };

    // does a complete update of an existing entity instance
    this.update = function (options, cb) {
        // handles optional 'options' argument
        if (!cb && typeof options === "function") {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};
        if (options.id === null || options.id === undefined) { return cb(new Error("'id' property is missing.")); }
        if (options.data === null || options.data === undefined) { return cb(new Error("'data' property is missing.")); }
        if (!options.resource || typeof options.resource !== "string") { return cb(new Error("'resource' property is missing or invalid.")); }

        self.oData({
            data: options.data,
            method: "PATCH",
            command: "/" + options.resource + "(" + buildODataId(options.id) + ")",
            etag: "*"
        }, cb);
    };

    // removes an existing instance from an entity set
    this.remove = function (options, cb) {
        // handles optional 'options' argument
        if (!cb && typeof options === "function") {
            cb = options;
            options = {};
        }

        // sets default values
        cb = cb || defaultCb;
        options = options || {};
        if (!options.resource || typeof options.resource !== "string") { return cb(new Error("'resource' property is missing or invalid.")); }
        if (options.id === null || options.id === undefined) { return cb(new Error("'id' property is missing.")); }

        self.oData({
            method: "DELETE",
            command: "/" + options.resource + "(" + buildODataId(options.id) + ")",
            etag: "*"
        }, cb);
    };

    // adds methods dynamically to the instance passed by parameter.
    // for each entity set, methods for get, update, create, etc. will be added.
    this._hook = function (target) {
        // function for add a single method
        function addMethod(method, prefix, entitySet) {

            target[prefix + entitySet] = function (options, cb) {

                if (!cb && typeof options === "function") {
                    cb = options;
                    options = {};
                }

                cb = cb || defaultCb;
                options = options || {};
                options.resource = entitySet;

                method(options, cb);
            };
        }

        // function for add all methods to a every entity set
        function addAllMethods(entitySetArray) {
            entitySetArray = entitySetArray || [];
            entitySetArray.forEach(function (entitySet) {
                addMethod(self.get, "get", entitySet);
                addMethod(self.query, "query", entitySet);
                addMethod(self.links, "links", entitySet);
                addMethod(self.count, "count", entitySet);
                addMethod(self.create, "create", entitySet);
                addMethod(self.replace, "replace", entitySet);
                addMethod(self.update, "update", entitySet);
                addMethod(self.remove, "remove", entitySet);
            });
        };

        // adds te methods if the array of entity sets was populated
        if (entitySets) {
            addAllMethods(entitySets);
        } else {
            // wait for the entitySets
            pendingHook = function () {
                addAllMethods(entitySets);
                pendingHook = null;
            };
        }
    };

    // returns
    this.lookupMethod = function (name, cb) {
        if (name==="_hook") return cb(); // ignorre hook method

        var method;

        if (!pendingHook) {
            method = self[name];
            if (typeof method === "function") return cb(null, method);
            return cb();
        }

        // returns a wrappers that waits until entity sets were retrieved
        function wrapper(options, callback) {

            // forces the invocation to entity sets
            self.entitySets(options, function (err) {
                if (err) { return callback(err); }

                // after entitySets were retrieved, invokes the method or returns not found (404)
                method = self[name];
                if (typeof method === "function") return method(options, callback);
                return callback(null, null);
            });
        };

        return cb(null, wrapper);
    };
};
