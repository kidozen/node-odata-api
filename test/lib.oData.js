"use strict";
var assert = require("assert");
var oData = null;

describe("lib.oData", function() {
    it("was able to require module", function() {
        oData = require("../lib/oData.js");
        assert.ok(oData);
    });

    describe("buildPath", function () {
        before(function () {
            oData = require("../lib/oData.js");
        });

        it("should replace // with / in the command", function () {
            var newPath = oData.buildPath("http://foo.com/srv_root", {command: "Entity//Property"});
            assert.equal(newPath, "http://foo.com/srv_root/Entity/Property");
        });

        it("should not duplicate / when href ends with /", function () {
            var newPath = oData.buildPath("http://foo.com/srv_root/", {command: "Entity//Property"});
            assert.equal(newPath, "http://foo.com/srv_root/Entity/Property");
        });
    });
});