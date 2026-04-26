try:
    from http.server import BaseHTTPRequestHandler, HTTPServer
except ImportError:
    from BaseHTTPServer import BaseHTTPRequestHandler, HTTPServer

HOST = "127.0.0.1"
PORT = 8080

A_INPUT = 201
A_OUTPUT = 501
D_INPUT = 4001
D_OUTPUT = 4001

MIN_POLL_MS = 50
DEFAULT_POLL_MS = 500

STATE = {
    "ain": [0] * A_INPUT,
    "aout": [0] * A_OUTPUT,
    "din": [0] * D_INPUT,
    "dout": [0] * D_OUTPUT,
}


def parse_query_string(path):
    query = {}

    if "?" not in path:
        return path, query

    route, qs = path.split("?", 1)

    pairs = qs.split("&")
    for pair in pairs:
        if "=" in pair:
            key, value = pair.split("=", 1)
            query[key] = value
        else:
            query[pair] = ""

    return route, query


def parse_post_body(body_text):
    data = {}

    if not body_text:
        return data

    pairs = body_text.split("&")
    for pair in pairs:
        if "=" in pair:
            key, value = pair.split("=", 1)
            data[key] = value
        else:
            data[pair] = ""

    return data


def html_escape(text):
    text = str(text)
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    text = text.replace('"', "&quot;")
    return text


def get_poll_ms(query):
    poll_ms_text = query.get("poll_ms", str(DEFAULT_POLL_MS))
    try:
        poll_ms = int(poll_ms_text)
    except:
        poll_ms = DEFAULT_POLL_MS

    if poll_ms < MIN_POLL_MS:
        poll_ms = MIN_POLL_MS

    return poll_ms


class SimpleHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        route, query = parse_query_string(self.path)

        if route == "/":
            self.send_text_response("hello from xp server")
            return

        if route == "/health":
            self.send_text_response("ok")
            return

        if route == "/data":
            self.send_text_response("test data from python server")
            return

        if route == "/read":
            self.handle_read(query)
            return

        if route == "/status":
            self.handle_status()
            return

        if route == "/ioio/status":
            # Combined text payload for VB: poll_ms, ain, aout, din, dout. The handler
            # picks the lines it cares about by prefix and ignores the rest.
            self.handle_ioio_status(query)
            return

        if route == "/api/analog":
            self.handle_api_analog(query)
            return

        if route == "/api/digital":
            self.handle_api_digital(query)
            return

        if route == "/test/analog":
            self.handle_test_analog(query)
            return

        if route == "/test/digital":
            self.handle_test_digital(query)
            return

        self.send_not_found()

    def do_POST(self):
        route, query = parse_query_string(self.path)

        content_length = self.headers.get("Content-Length")
        if content_length is None:
            length = 0
        else:
            length = int(content_length)

        raw_body = self.rfile.read(length)

        try:
            body_text = raw_body.decode("utf-8")
        except:
            body_text = raw_body

        form = parse_post_body(body_text)

        for key in query:
            if key not in form:
                form[key] = query[key]

        if route == "/write":
            self.handle_write(form)
            return

        if route == "/api/set_analog":
            self.handle_set_analog(form)
            return

        if route == "/api/set_digital":
            self.handle_set_digital(form)
            return

        self.send_not_found()

    def handle_read(self, query):
        array_name = query.get("array", "")
        index_text = query.get("index", "")

        if array_name not in STATE:
            self.send_text_response("error: invalid array", 400)
            return

        try:
            index = int(index_text)
        except:
            self.send_text_response("error: invalid index", 400)
            return

        max_len = len(STATE[array_name])

        if index < 0 or index >= max_len:
            self.send_text_response("error: index out of range", 400)
            return

        value = STATE[array_name][index]

        print("READ  array=%s index=%d value=%s" % (array_name, index, value))
        self.send_text_response(str(value), 200)

    def handle_write(self, form):
        array_name = form.get("array", "")
        index_text = form.get("index", "")
        value_text = form.get("value", "")

        if array_name not in STATE:
            self.send_text_response("error: invalid array", 400)
            return

        try:
            index = int(index_text)
        except:
            self.send_text_response("error: invalid index", 400)
            return

        max_len = len(STATE[array_name])

        if index < 0 or index >= max_len:
            self.send_text_response("error: index out of range", 400)
            return

        try:
            value = int(value_text)
        except:
            self.send_text_response("error: invalid value", 400)
            return

        STATE[array_name][index] = value

        print("WRITE array=%s index=%d value=%d" % (array_name, index, value))
        self.send_text_response("ok", 200)

    def handle_set_analog(self, form):
        array_name = form.get("array", "")
        index_text = form.get("index", "")
        value_text = form.get("value", "")

        if array_name not in ["ain", "aout"]:
            self.send_text_response("error: invalid analog array", 400)
            return

        try:
            index = int(index_text)
            value = int(value_text)
        except:
            self.send_text_response("error: invalid parameter", 400)
            return

        if index < 0 or index >= len(STATE[array_name]):
            self.send_text_response("error: index out of range", 400)
            return

        if value < 0:
            value = 0
        if value > 255:
            value = 255

        STATE[array_name][index] = value
        print("SET ANALOG array=%s index=%d value=%d" % (array_name, index, value))
        self.send_text_response("ok", 200)

    def handle_set_digital(self, form):
        array_name = form.get("array", "")
        index_text = form.get("index", "")
        value_text = form.get("value", "")

        if array_name not in ["din", "dout"]:
            self.send_text_response("error: invalid digital array", 400)
            return

        try:
            index = int(index_text)
            value = int(value_text)
        except:
            self.send_text_response("error: invalid parameter", 400)
            return

        if index < 0 or index >= len(STATE[array_name]):
            self.send_text_response("error: index out of range", 400)
            return

        if array_name == "din":
            if value <= 0:
                value = 0
            else:
                value = 1
        else:
            if value <= 0:
                value = 0
            elif value == 1:
                value = 1
            else:
                value = 2

        STATE[array_name][index] = value
        print("SET DIGITAL array=%s index=%d value=%d" % (array_name, index, value))
        self.send_text_response("ok", 200)

    def handle_api_analog(self, query):
        lines = []
        lines.append("poll_ms=%d" % get_poll_ms(query))
        lines.append("ain=" + ",".join([str(x) for x in STATE["ain"]]))
        lines.append("aout=" + ",".join([str(x) for x in STATE["aout"]]))
        self.send_text_response("\n".join(lines), 200)

    def handle_ioio_status(self, query):
        lines = []
        lines.append("poll_ms=%d" % get_poll_ms(query))
        lines.append("ain=" + ",".join([str(x) for x in STATE["ain"]]))
        lines.append("aout=" + ",".join([str(x) for x in STATE["aout"]]))
        lines.append("din=" + ",".join([str(x) for x in STATE["din"]]))
        lines.append("dout=" + ",".join([str(x) for x in STATE["dout"]]))
        self.send_text_response("\n".join(lines), 200)

    def handle_api_digital(self, query):
        lines = []
        lines.append("poll_ms=%d" % get_poll_ms(query))
        lines.append("din=" + ",".join([str(x) for x in STATE["din"]]))
        lines.append("dout=" + ",".join([str(x) for x in STATE["dout"]]))
        self.send_text_response("\n".join(lines), 200)

    def handle_status(self):
        html = self.build_status_html()
        self.send_html_response(html, 200)

    def handle_test_analog(self, query):
        poll_ms = get_poll_ms(query)
        html = self.build_test_analog_html(poll_ms)
        self.send_html_response(html, 200)

    def handle_test_digital(self, query):
        poll_ms = get_poll_ms(query)
        html = self.build_test_digital_html(poll_ms)
        self.send_html_response(html, 200)

    def build_status_html(self):
        parts = []

        parts.append("<html>")
        parts.append("<head>")
        parts.append("<title>XP Python Server Status</title>")
        parts.append("<style>")
        parts.append("body { font-family: Arial, sans-serif; font-size: 12px; margin: 12px; }")
        parts.append("h1 { margin-bottom: 4px; }")
        parts.append(".toolbar { margin-bottom: 10px; padding: 6px; border: 1px solid #999999; background: #f0f0f0; }")
        parts.append(".col { float: left; width: 24%; margin-right: 1%; border: 1px solid #999999; }")
        parts.append(".col h2 { margin: 0; padding: 6px; background: #dddddd; font-size: 14px; }")
        parts.append(".list { height: 500px; overflow: auto; }")
        parts.append(".row { padding: 3px 6px; border-top: 1px solid #eeeeee; white-space: nowrap; }")
        parts.append(".idx { display: inline-block; width: 45px; color: #666666; }")
        parts.append(".val { display: inline-block; width: 80px; font-weight: bold; }")
        parts.append(".footer { clear: both; padding-top: 12px; color: #666666; }")
        parts.append("</style>")

        parts.append("<script type='text/javascript'>")
        parts.append("var POLL_MS = %d;" % DEFAULT_POLL_MS)
        parts.append("var MIN_POLL_MS = %d;" % MIN_POLL_MS)
        parts.append("var isUpdatingStatus = false;")

        parts.append("""
    function makeRequest(method, url, callback) {
        var xhr = new XMLHttpRequest();
        xhr.onreadystatechange = function() {
            if (xhr.readyState == 4) {
                callback(xhr);
            }
        };
        xhr.open(method, url, true);
        xhr.send(null);
    }

    function updateStatusValue(arrayName, index, value) {
        var el = document.getElementById(arrayName + '_value_' + index);
        if (el && el.innerHTML != String(value)) {
            el.innerHTML = value;
        }
    }

    function applyStatusArray(arrayName, csv) {
        var values = csv.split(',');
        var i;
        for (i = 0; i < values.length; i++) {
            updateStatusValue(arrayName, i, parseInt(values[i], 10));
        }
    }

    function parseStatusAnalog(text) {
        var lines = text.split('\\n');
        var i;
        for (i = 0; i < lines.length; i++) {
            if (lines[i].indexOf('ain=') == 0) {
                applyStatusArray('ain', lines[i].substring(4));
            }
            else if (lines[i].indexOf('aout=') == 0) {
                applyStatusArray('aout', lines[i].substring(5));
            }
        }
    }

    function parseStatusDigital(text) {
        var lines = text.split('\\n');
        var i;
        for (i = 0; i < lines.length; i++) {
            if (lines[i].indexOf('din=') == 0) {
                applyStatusArray('din', lines[i].substring(4));
            }
            else if (lines[i].indexOf('dout=') == 0) {
                applyStatusArray('dout', lines[i].substring(5));
            }
        }
    }

    function pollStatus() {
        if (isUpdatingStatus) {
            return;
        }

        isUpdatingStatus = true;

        makeRequest('GET', '/api/analog?poll_ms=' + POLL_MS, function(xhr1) {
            if (xhr1.status == 200) {
                parseStatusAnalog(xhr1.responseText);
            }

            makeRequest('GET', '/api/digital?poll_ms=' + POLL_MS, function(xhr2) {
                isUpdatingStatus = false;
                if (xhr2.status == 200) {
                    parseStatusDigital(xhr2.responseText);
                }
            });
        });
    }

    function changeStatusPollMs() {
        var box = document.getElementById('poll_ms_box');
        var value = parseInt(box.value, 10);

        if (isNaN(value) || value < MIN_POLL_MS) {
            value = MIN_POLL_MS;
        }

        POLL_MS = value;
        box.value = value;

        window.clearInterval(window._statusTimer);
        window._statusTimer = window.setInterval(pollStatus, POLL_MS);
    }

    function initStatusPage() {
        pollStatus();
        window._statusTimer = window.setInterval(pollStatus, POLL_MS);
    }
    """)
        parts.append("</script>")
        parts.append("</head>")
        parts.append("<body onload='initStatusPage();'>")
        parts.append("<h1>XP Python Server Status</h1>")
        parts.append("<div class='toolbar'>")
        parts.append("Poll ms: <input type='text' id='poll_ms_box' value='%d' size='6'>" % DEFAULT_POLL_MS)
        parts.append(" <input type='button' value='Apply Poll Rate' onclick='changeStatusPollMs();'>")
        parts.append(" Minimum: %d ms" % MIN_POLL_MS)
        parts.append("</div>")

        array_order = ["ain", "aout", "din", "dout"]
        array_limits = {
            "ain": len(STATE["ain"]),
            "aout": len(STATE["aout"]),
            "din": min(len(STATE["din"]), 200),
            "dout": min(len(STATE["dout"]), 200),
        }

        for name in array_order:
            parts.append("<div class='col'>")
            parts.append("<h2>%s</h2>" % name)
            parts.append("<div class='list'>")

            i = 0
            while i < array_limits[name]:
                parts.append(
                    "<div class='row'><span class='idx'>[%d]</span><span class='val' id='%s_value_%d'>%s</span></div>"
                    % (i, name, i, STATE[name][i])
                )
                i += 1

            if array_limits[name] < len(STATE[name]):
                parts.append("<div class='row'>... (%d more not shown)</div>" % (len(STATE[name]) - array_limits[name]))

            parts.append("</div>")
            parts.append("</div>")

        parts.append("<div class='footer'>")
        parts.append("ain length: %d<br>" % len(STATE["ain"]))
        parts.append("aout length: %d<br>" % len(STATE["aout"]))
        parts.append("din length: %d<br>" % len(STATE["din"]))
        parts.append("dout length: %d<br>" % len(STATE["dout"]))
        parts.append("</div>")

        parts.append("</body>")
        parts.append("</html>")

        return "".join(parts)

    def build_test_analog_html(self, poll_ms):
        parts = []

        parts.append("<html>")
        parts.append("<head>")
        parts.append("<title>Analog I/O Test Page</title>")
        parts.append("<style>")
        parts.append("body { font-family: Arial, sans-serif; font-size: 12px; margin: 10px; }")
        parts.append("h1 { margin-bottom: 4px; }")
        parts.append(".toolbar { margin-bottom: 10px; padding: 6px; border: 1px solid #999999; background: #f0f0f0; }")
        parts.append(".col { float: left; width: 48%; margin-right: 1%; }")
        parts.append(".panel { border: 1px solid #999999; }")
        parts.append(".panel h2 { margin: 0; padding: 6px; background: #dddddd; }")
        parts.append(".list { height: 500px; overflow: auto; }")
        parts.append(".row { padding: 4px; border-top: 1px solid #eeeeee; white-space: nowrap; }")
        parts.append(".idx { display: inline-block; width: 45px; }")
        parts.append(".valbox { display: inline-block; width: 40px; text-align: right; margin-left: 6px; }")
        parts.append("</style>")
        parts.append("<script type='text/javascript'>")
        parts.append("var POLL_MS = %d;" % poll_ms)
        parts.append("var MIN_POLL_MS = %d;" % MIN_POLL_MS)
        parts.append("var isUpdating = false;")
        parts.append("""
function makeRequest(method, url, callback) {
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function() {
        if (xhr.readyState == 4) {
            callback(xhr);
        }
    };
    xhr.open(method, url, true);
    xhr.send(null);
}

function setAnalogValue(arrayName, index, value) {
    var url = '/api/set_analog?array=' + escape(arrayName) + '&index=' + index + '&value=' + value;
    makeRequest('POST', url, function(xhr) {});
}

function updateSliderAndLabel(arrayName, index, value) {
    var slider = document.getElementById(arrayName + '_slider_' + index);
    var label = document.getElementById(arrayName + '_value_' + index);

    if (slider) {
        if (document.activeElement != slider) {
            if (parseInt(slider.value, 10) != value) {
                slider.value = value;
            }
        }
    }

    if (label) {
        if (label.innerHTML != String(value)) {
            label.innerHTML = value;
        }
    }
}

function parseAnalogResponse(text) {
    var lines = text.split('\\n');
    var i;
    for (i = 0; i < lines.length; i++) {
        if (lines[i].indexOf('ain=') == 0) {
            applyArrayValues('ain', lines[i].substring(4));
        }
        else if (lines[i].indexOf('aout=') == 0) {
            applyArrayValues('aout', lines[i].substring(5));
        }
    }
}

function applyArrayValues(arrayName, csv) {
    var values = csv.split(',');
    var i;
    for (i = 0; i < values.length; i++) {
        updateSliderAndLabel(arrayName, i, parseInt(values[i], 10));
    }
}

function pollAnalog() {
    if (isUpdating) {
        return;
    }
    isUpdating = true;

    makeRequest('GET', '/api/analog?poll_ms=' + POLL_MS, function(xhr) {
        isUpdating = false;
        if (xhr.status == 200) {
            parseAnalogResponse(xhr.responseText);
        }
    });
}

function onSliderChange(arrayName, index, sliderObj) {
    var value = parseInt(sliderObj.value, 10);
    updateSliderAndLabel(arrayName, index, value);
    setAnalogValue(arrayName, index, value);
}

function changePollMs() {
    var box = document.getElementById('poll_ms_box');
    var value = parseInt(box.value, 10);
    if (isNaN(value) || value < MIN_POLL_MS) {
        value = MIN_POLL_MS;
    }
    POLL_MS = value;
    box.value = value;
    window.clearInterval(window._analogTimer);
    window._analogTimer = window.setInterval(pollAnalog, POLL_MS);
}

function initAnalogPage() {
    pollAnalog();
    window._analogTimer = window.setInterval(pollAnalog, POLL_MS);
}
""")
        parts.append("</script>")
        parts.append("</head>")
        parts.append("<body onload='initAnalogPage();'>")
        parts.append("<h1>Analog I/O Test Page</h1>")
        parts.append("<div class='toolbar'>")
        parts.append("Poll ms: <input type='text' id='poll_ms_box' value='%d' size='6'>" % poll_ms)
        parts.append(" <input type='button' value='Apply Poll Rate' onclick='changePollMs();'>")
        parts.append(" Minimum: %d ms" % MIN_POLL_MS)
        parts.append("</div>")

        parts.append("<div class='col'>")
        parts.append("<div class='panel'>")
        parts.append("<h2>Analog Inputs (ain)</h2>")
        parts.append("<div class='list'>")
        i = 0
        while i < len(STATE["ain"]):
            parts.append("<div class='row'>")
            parts.append("<span class='idx'>[%d]</span>" % i)
            parts.append("<input type='range' min='0' max='255' value='0' id='ain_slider_%d' onchange=\"onSliderChange('ain', %d, this);\" />" % (i, i))
            parts.append("<span class='valbox' id='ain_value_%d'>0</span>" % i)
            parts.append("</div>")
            i += 1
        parts.append("</div>")
        parts.append("</div>")
        parts.append("</div>")

        parts.append("<div class='col'>")
        parts.append("<div class='panel'>")
        parts.append("<h2>Analog Outputs (aout)</h2>")
        parts.append("<div class='list'>")
        i = 0
        while i < len(STATE["aout"]):
            parts.append("<div class='row'>")
            parts.append("<span class='idx'>[%d]</span>" % i)
            parts.append("<input type='range' min='0' max='255' value='0' id='aout_slider_%d' onchange=\"onSliderChange('aout', %d, this);\" />" % (i, i))
            parts.append("<span class='valbox' id='aout_value_%d'>0</span>" % i)
            parts.append("</div>")
            i += 1
        parts.append("</div>")
        parts.append("</div>")
        parts.append("</div>")

        parts.append("</body>")
        parts.append("</html>")

        return "".join(parts)

    def build_test_digital_html(self, poll_ms):
        parts = []

        parts.append("<html>")
        parts.append("<head>")
        parts.append("<title>Digital I/O Test Page</title>")
        parts.append("<style>")
        parts.append("body { font-family: Arial, sans-serif; font-size: 12px; margin: 10px; }")
        parts.append("h1 { margin-bottom: 4px; }")
        parts.append(".toolbar { margin-bottom: 10px; padding: 6px; border: 1px solid #999999; background: #f0f0f0; }")
        parts.append(".col { float: left; width: 48%; margin-right: 1%; }")
        parts.append(".panel { border: 1px solid #999999; }")
        parts.append(".panel h2 { margin: 0; padding: 6px; background: #dddddd; }")
        parts.append(".list { height: 500px; overflow: auto; }")
        parts.append(".row { padding: 4px; border-top: 1px solid #eeeeee; white-space: nowrap; }")
        parts.append(".idx { display: inline-block; width: 55px; }")
        parts.append(".statebox { display: inline-block; width: 60px; margin-left: 8px; font-weight: bold; }")
        parts.append("</style>")
        parts.append("<script type='text/javascript'>")
        parts.append("var POLL_MS = %d;" % poll_ms)
        parts.append("var MIN_POLL_MS = %d;" % MIN_POLL_MS)
        parts.append("var isUpdating = false;")
        parts.append("""
function makeRequest(method, url, callback) {
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function() {
        if (xhr.readyState == 4) {
            callback(xhr);
        }
    };
    xhr.open(method, url, true);
    xhr.send(null);
}

function getDigitalLabel(arrayName, value) {
    if (arrayName == 'din') {
        if (value <= 0) {
            return 'Off';
        }
        return 'On';
    }

    if (value <= 0) {
        return 'Off';
    }
    if (value == 1) {
        return 'On';
    }
    return 'Flash';
}

function setDigitalValue(arrayName, index, value) {
    var url = '/api/set_digital?array=' + escape(arrayName) + '&index=' + index + '&value=' + value;
    makeRequest('POST', url, function(xhr) {});
}

function updateDigitalControl(arrayName, index, value) {
    var sel = document.getElementById(arrayName + '_select_' + index);
    var label = document.getElementById(arrayName + '_value_' + index);
    var normalizedValue;
    var labelText;

    if (arrayName == 'din') {
        if (value <= 0) {
            normalizedValue = '0';
            labelText = 'Off';
        } else {
            normalizedValue = '1';
            labelText = 'On';
        }
    } else {
        if (value <= 0) {
            normalizedValue = '0';
            labelText = 'Off';
        } else if (value == 1) {
            normalizedValue = '1';
            labelText = 'On';
        } else {
            normalizedValue = '2';
            labelText = 'Flash';
        }
    }

    if (sel) {
        if (document.activeElement != sel) {
            if (sel.value != normalizedValue) {
                sel.value = normalizedValue;
            }
        }
    }

    if (label) {
        if (label.innerHTML != labelText) {
            label.innerHTML = labelText;
        }
    }
}
                     
function isDigitalControlActive() {
    var el = document.activeElement;
    if (!el) {
        return false;
    }

    if (el.tagName && el.tagName.toUpperCase() == 'SELECT') {
        return true;
    }

    return false;
}

function applyDigitalArray(arrayName, csv) {
    var values = csv.split(',');
    var i;
    for (i = 0; i < values.length; i++) {
        updateDigitalControl(arrayName, i, parseInt(values[i], 10));
    }
}

function parseDigitalResponse(text) {
    var lines = text.split('\\n');
    var i;
    for (i = 0; i < lines.length; i++) {
        if (lines[i].indexOf('din=') == 0) {
            applyDigitalArray('din', lines[i].substring(4));
        }
        else if (lines[i].indexOf('dout=') == 0) {
            applyDigitalArray('dout', lines[i].substring(5));
        }
    }
}

function pollDigital() {
    if (isUpdating) {
        return;
    }

    // Avoid repaint flicker while user is interacting with a dropdown
    if (isDigitalControlActive()) {
        return;
    }

    isUpdating = true;

    makeRequest('GET', '/api/digital?poll_ms=' + POLL_MS, function(xhr) {
        isUpdating = false;
        if (xhr.status == 200) {
            parseDigitalResponse(xhr.responseText);
        }
    });
}

function onDigitalChange(arrayName, index, selObj) {
    var value = parseInt(selObj.value, 10);
    updateDigitalControl(arrayName, index, value);
    setDigitalValue(arrayName, index, value);
}

function changePollMs() {
    var box = document.getElementById('poll_ms_box');
    var value = parseInt(box.value, 10);
    if (isNaN(value) || value < MIN_POLL_MS) {
        value = MIN_POLL_MS;
    }
    POLL_MS = value;
    box.value = value;
    window.clearInterval(window._digitalTimer);
    window._digitalTimer = window.setInterval(pollDigital, POLL_MS);
}

function initDigitalPage() {
    pollDigital();
    window._digitalTimer = window.setInterval(pollDigital, POLL_MS);
}
""")
        parts.append("</script>")
        parts.append("</head>")
        parts.append("<body onload='initDigitalPage();'>")
        parts.append("<h1>Digital I/O Test Page</h1>")
        parts.append("<div class='toolbar'>")
        parts.append("Poll ms: <input type='text' id='poll_ms_box' value='%d' size='6'>" % poll_ms)
        parts.append(" <input type='button' value='Apply Poll Rate' onclick='changePollMs();'>")
        parts.append(" Minimum: %d ms" % MIN_POLL_MS)
        parts.append("</div>")

        parts.append("<div class='col'>")
        parts.append("<div class='panel'>")
        parts.append("<h2>Digital Inputs (din)</h2>")
        parts.append("<div class='list'>")
        i = 0
        max_din_display = min(len(STATE["din"]), 200)
        parts.append("<div class='row'>Showing first %d of %d</div>" % (max_din_display, len(STATE["din"]))) 
        while i < max_din_display:
            parts.append("<div class='row'>")
            parts.append("<span class='idx'>[%d]</span>" % i)
            parts.append("<select id='din_select_%d' onchange=\"onDigitalChange('din', %d, this);\">" % (i, i))
            parts.append("<option value='0'>Off</option>")
            parts.append("<option value='1'>On</option>")
            parts.append("</select>")
            parts.append("<span class='statebox' id='din_value_%d'>Off</span>" % i)
            parts.append("</div>")
            i += 1
        parts.append("</div>")
        parts.append("</div>")
        parts.append("</div>")

        parts.append("<div class='col'>")
        parts.append("<div class='panel'>")
        parts.append("<h2>Digital Outputs (dout)</h2>")
        parts.append("<div class='list'>")
        i = 0
        max_dout_display = min(len(STATE["dout"]), 200)
        parts.append("<div class='row'>Showing first %d of %d</div>" % (max_dout_display, len(STATE["dout"])))
        while i < max_dout_display:
            parts.append("<div class='row'>")
            parts.append("<span class='idx'>[%d]</span>" % i)
            parts.append("<select id='dout_select_%d' onchange=\"onDigitalChange('dout', %d, this);\">" % (i, i))
            parts.append("<option value='0'>Off</option>")
            parts.append("<option value='1'>On</option>")
            parts.append("<option value='2'>Flash</option>")
            parts.append("</select>")
            parts.append("<span class='statebox' id='dout_value_%d'>Off</span>" % i)
            parts.append("</div>")
            i += 1
        parts.append("</div>")
        parts.append("</div>")
        parts.append("</div>")

        parts.append("</body>")
        parts.append("</html>")

        return "".join(parts)

    def send_text_response(self, text, status_code=200):
        body = self.to_bytes(text)
        self.send_response(status_code)
        self.send_header("Content-Type", "text/plain")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_html_response(self, html, status_code=200):
        body = self.to_bytes(html)
        self.send_response(status_code)
        self.send_header("Content-Type", "text/html")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_not_found(self):
        self.send_text_response("not found", 404)

    def log_message(self, format, *args):
        try:
            print("[%s] %s" % (self.address_string(), format % args))
        except:
            pass

    def to_bytes(self, text):
        try:
            return text.encode("utf-8")
        except:
            return text


def main():
    print("Starting Python HTTP server on http://%s:%s" % (HOST, PORT))
    print("Routes:")
    print("  GET  /")
    print("  GET  /health")
    print("  GET  /data")
    print("  GET  /status")
    print("  GET  /test/analog")
    print("  GET  /test/digital")
    print("  GET  /api/analog")
    print("  GET  /api/digital")
    print("  GET  /ioio/status   (text alias of /api/analog for VB6 client)")
    print("  POST /api/set_analog?array=ain&index=0&value=123")
    print("  POST /api/set_digital?array=dout&index=0&value=2")
    print("Array lengths:")
    print("  ain  =", A_INPUT)
    print("  aout =", A_OUTPUT)
    print("  din  =", D_INPUT)
    print("  dout =", D_OUTPUT)

    server = HTTPServer((HOST, PORT), SimpleHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
