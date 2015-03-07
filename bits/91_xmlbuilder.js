/////////////////////////////////////////////////////////////////////////////////////////////////////
var XmlNode = (function () {
  function XmlNode(tagName, attributes, children) {

    if (!(this instanceof XmlNode)) {
      return new XmlNode(tagName, attributes, children);
    }
    this.tagName = tagName;
    this._attributes = attributes || {};
    this._children = children || [];
    this._prefix = '';
    return this;
  }

  XmlNode.prototype.createElement = function () {
    return new XmlNode(arguments)
  }

  XmlNode.prototype.children = function() {
    return this._children;
  }

  XmlNode.prototype.append = function (node) {
    this._children.push(node);
    return this;
  }

  XmlNode.prototype.prefix = function (prefix) {
    if (arguments.length==0) { return this._prefix;}
    this._prefix = prefix;
    return this;
  }

  XmlNode.prototype.attr = function (attr, value) {
    if (value == undefined) {
      delete this._attributes[attr];
      return this;
    }
    if (arguments.length == 0) {
      return this._attributes;
    }
    else if (typeof attr == 'string' && arguments.length == 1) {
      return this._attributes.attr[attr];
    }
    if (typeof attr == 'object' && arguments.length == 1) {
      for (var key in attr) {
        this._attributes[key] = attr[key];
      }
    }
    else if (arguments.length == 2 && typeof attr == 'string') {
      this._attributes[attr] = value;
    }
    return this;
  }

  var APOS = "'"; QUOTE = '"'
  var ESCAPED_QUOTE = {  }
  ESCAPED_QUOTE[QUOTE] = '&quot;'
  ESCAPED_QUOTE[APOS] = '&apos;'

  XmlNode.prototype.escapeAttributeValue = function(att_value) {
    return '"' + att_value.replace(/\"/g,'&quot;') + '"';// TODO Extend with four other codes

  }

  XmlNode.prototype.toXml = function (node) {
    if (!node) node = this;
    var xml = node._prefix;
    xml += '<' + node.tagName;
    if (node._attributes) {
      for (var key in node._attributes) {
        xml += ' ' + key + '=' + this.escapeAttributeValue(''+node._attributes[key]) + ''
      }
    }
    if (node._children && node._children.length > 0) {
      xml += ">";
      for (var i = 0; i < node._children.length; i++) {
        xml += this.toXml(node._children[i]);
      }
      xml += '</' + node.tagName + '>';
    }
    else {
      xml += '/>';
    }
    return xml;
  }
  return XmlNode;
})();
