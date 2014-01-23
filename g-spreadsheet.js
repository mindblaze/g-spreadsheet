
var xml2json = require('xml2json');
var request = require('request');
var http = require('http');
var querystring = require('querystring');

var BASE_URL = "https://spreadsheets.google.com/feeds/";

function doFeedRequest(spreadSheetId, workSheetId, oauth, type, query, cb) {
  var requestURL = BASE_URL + type + '/' + spreadSheetId + '/' + (workSheetId? workSheetId + '/':'') + (oauth? 'private' : 'public') + '/full' + (query? '?'+querystring.stringify(query):'');
  request.get(
    {
      url: requestURL,
      headers: {Authorization: 'Bearer ' + oauth.token}
    }, function (e, req, body) {
      if (e) return cb(e);
      switch (req.statusCode) {
        case 200:
          return cb(null, JSON.parse(xml2json.toJson(body)).feed);
        default:
          return cb(req.statusCode + ' ' + http.STATUS_CODES[req.statusCode]);
      }
    }
  );
}

function GoogleSpreadSheet(spreadSheetId, oauth) {
  
  var self = this;
  this.id = spreadSheetId;
  this.updated = null;
  this.author = {};
  this.title = null;
  this.worksheets = [];

  this.getInfo = function(query, cb) {
    if (typeof query == 'function') {
      cb = query;
      query = null;
    }
    doFeedRequest(self.id, null, oauth, 'worksheets', query, function (e, body) {
      self.title = body.title.$t;
      self.updated = body.updated;
      self.author = body.author;
      var entries = body.entry;
      if (!Array.isArray(body.entry)) entries = [body.entry];
      for (var i = 0; i < entries.length; i++) {
        var current = entries[i];
        var newEntry = {};
        newEntry.id = current.id.match(/\/([a-zA-z0-9\-\_]+)$/)[1];
        newEntry.title = current.title.$t;
        newEntry.rows = current['gs:rowCount'];
        newEntry.cols = current['gs:colCount'];
        self.worksheets.push(new WorkSheet(self.id, newEntry.id, oauth, newEntry.title, newEntry.rows, newEntry.cols));
      }
      return cb(null, self);
    });
  }

}

function WorkSheet(spreadSheetId, workSheetId, oauth, title, rows, cols) {
  var self = this;
  this.sId = spreadSheetId;
  this.id = workSheetId;
  this.title = title;
  this.rows = rows;
  this.cols = cols;

  this.getRows = function (query, cb) {
    if (typeof query == 'function') {
      cb = query;
      query = null;
    }
    var result = [];
    this.getCells({"min-row": 1, "max-row": 1, "min-col": 1, "max-col": self.cols}, function (e, body) {
      if (e) return cb(e);
      var columnHeader = [];
      var entries = body.entry;
      if (!Array.isArray(body.entry)) entries = [body.entry];
      entries.forEach(function(entry) {
        columnHeader.push(entry['gs:cell']['$t']);
      });
      doFeedRequest(self.sId, self.id, oauth, 'list', query, function (e, body) {
        if (e) return cb(e);
        var entries = body.entry;
        if (!Array.isArray(body.entry)) entries = [body.entry];
        for (var i = 0; i < entries.length; i++) {
          var entry = entries[i];
          var newEntry = {};
          for (var j = 0; j < columnHeader.length; j++) {
            var rowKey = columnHeader[j].replace(/[^a-zA-Z0-9]+/g, '').toLowerCase();
            var value = entry['gsx:'+rowKey];
            if (typeof value == 'object' && Object.keys(value).length == 0) value = null;
            else if (value == undefined) value = null;
            newEntry[columnHeader[j]] = value;
          }
          result.push(newEntry);
        }
        cb(null, result);
      });
    });
  }

  this.getCells = function (query, cb) {
    if (typeof query == 'function') {
      cb = query;
      query = null;
    }

    doFeedRequest(self.sId, self.id, oauth, 'cells', query, function (e, body) {
      if (e) return cb(e);
      return cb(null, body);
    });
  }

}

if (typeof module == 'object')
  module.exports = GoogleSpreadSheet;