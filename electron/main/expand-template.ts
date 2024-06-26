export const expandTemplate = (opts: any) => {
  var sep = opts ? opts.sep : "{}";
  var len = sep.length;
  var whitespace = "\\s*";
  var left = escape(sep.substring(0, len / 2)) + whitespace;
  var right = whitespace + escape(sep.substring(len / 2, len));
  return function(template: any, values: any) {
    Object.keys(values).forEach(function(key) {
      var value = String(values[key]).replace(/\$/g, "$$$$");
      template = template.replace(regExp(key), value);
    });
    return template;
  };
  function escape(s: any) {
    return [].map.call(s, function(char) {
      return "\\" + char;
    }).join("");
  }
  function regExp(key: any) {
    return new RegExp(left + key + right, "g");
  }
}