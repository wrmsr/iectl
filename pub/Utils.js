//include('json2.js');

function click(id){
	puts('Click: ' + id);
	document.getElementById(id).click();
}

function getVal(id){
	puts('GetVal: ' + id);
	return document.getElementById(id).value;
}

function setVal(id, val){
	puts('SetVal: ' + id + ' = \"' + val + '\"');
	document.getElementById(id).value = val;
}

function am(){
	ie.Activate();
	ie.Max();
}

function $(fn){
	return reval('return (' + fn + ')();');
}

function trim(str){
	return str.replace(/^\s*/, '').replace(/\s*$/, '');
}

function wait(){
  ie.Wait();
}

function jq(){
  wait();
	reval(read('jquery-1.4.4.min.js'));
}

function ddSel(id, txt){
  return reval(
    'var dd = $(\'#' + id + '\'); ' +
    'var val = dd.children(\'option:contains("Biobank - QT")\').val(); ' +
    'if(!val) return false; ' +
    'dd.val(val); ' +
    'var cc = dd.get(0).onchange(); ' +
    'if(cc) cc(); ' +
    'return true;');
}
