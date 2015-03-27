"use strict";
var assert = require("assert");
var oData = null;

describe("lib.oData", function() {
    it("was able to require module", function() {
        oData = require("../lib/oData.js");
        assert.ok(oData);
    });
});