import requests

# Full HTTP status code descriptions
status_descriptions = {
    100: "Continue – Server received headers, continue with body.",
    101: "Switching Protocols – Switching to a different protocol.",
    102: "Processing – WebDAV: server is processing the request.",
    103: "Early Hints – Hints for preloading resources.",

    200: "OK – Request was successful.",
    201: "Created – Resource successfully created.",
    202: "Accepted – Request accepted, processing not complete.",
    203: "Non-Authoritative Information – Info from 3rd-party source.",
    204: "No Content – Success, no content returned.",
    205: "Reset Content – Client should reset document view.",
    206: "Partial Content – Partial GET request successful.",
    207: "Multi-Status – WebDAV: multiple statuses.",
    208: "Already Reported – WebDAV: element already reported.",
    226: "IM Used – Response includes instance manipulation.",

    300: "Multiple Choices – Multiple options available.",
    301: "Moved Permanently – Resource permanently moved.",
    302: "Found – Resource temporarily moved.",
    303: "See Other – Use GET to retrieve resource.",
    304: "Not Modified – Cached version is valid.",
    305: "Use Proxy – Deprecated.",
    306: "Switch Proxy – Unused/reserved.",
    307: "Temporary Redirect – Use same method for redirection.",
    308: "Permanent Redirect – Use same method for permanent move.",

    400: "Bad Request – Invalid syntax.",
    401: "Unauthorized – Authentication required.",
    402: "Payment Required – Reserved for future use.",
    403: "Forbidden – You don't have permission.",
    404: "Not Found – The requested URL doesn't exist.",
    405: "Method Not Allowed – Method not supported.",
    406: "Not Acceptable – Cannot fulfill request.",
    407: "Proxy Authentication Required – Auth needed from proxy.",
    408: "Request Timeout – Server timed out.",
    409: "Conflict – Request conflicts with server state.",
    410: "Gone – Resource is gone permanently.",
    411: "Length Required – Content-Length header missing.",
    412: "Precondition Failed – Preconditions not met.",
    413: "Payload Too Large – Entity too big.",
    414: "URI Too Long – URL too long.",
    415: "Unsupported Media Type – Format not supported.",
    416: "Range Not Satisfiable – Invalid range request.",
    417: "Expectation Failed – Expectation header cannot be met.",
    418: "I'm a teapot – Joke from RFC 2324.",
    421: "Misdirected Request – Sent to wrong server.",
    422: "Unprocessable Entity – WebDAV: can't process instructions.",
    423: "Locked – WebDAV: resource is locked.",
    424: "Failed Dependency – WebDAV: previous request failed.",
    425: "Too Early – Request replayed too soon.",
    426: "Upgrade Required – Must upgrade protocol.",
    428: "Precondition Required – Must supply conditions.",
    429: "Too Many Requests – Rate limited.",
    431: "Request Header Fields Too Large – Headers too large.",
    451: "Unavailable For Legal Reasons – Blocked due to law.",

    500: "Internal Server Error – Generic server failure.",
    501: "Not Implemented – Functionality not supported.",
    502: "Bad Gateway – Invalid response from upstream.",
    503: "Service Unavailable – Server overloaded or down.",
    504: "Gateway Timeout – Upstream server did not respond.",
    505: "HTTP Version Not Supported – Unsupported HTTP version.",
    506: "Variant Also Negotiates – Content negotiation error.",
    507: "Insufficient Storage – WebDAV: not enough storage.",
    508: "Loop Detected – Infinite loop in processing.",
    510: "Not Extended – More extensions required.",
    511: "Network Authentication Required – Needs auth to access."
}


def check_url_status():
    url = input("Enter a URL (e.g., https://example.com): ").strip()

    # Automatically add https:// if scheme is missing
    if not url.startswith("http://") and not url.startswith("https://"):
        url = "https://" + url

    try:
        response = requests.get(url)
        code = response.status_code
        reason = response.reason
        description = status_descriptions.get(
            code, "No detailed explanation available.")

        print(f"\nURL: {url}")
        print(f"Status Code: {code} ({reason})")
        print(f"Description: {description}")

    except requests.exceptions.MissingSchema:
        print("Invalid URL format. Include http:// or https://")
    except requests.exceptions.ConnectionError:
        print("Connection Error. The server could not be reached.")
    except requests.exceptions.Timeout:
        print("Request Timed Out.")
    except requests.exceptions.RequestException as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    check_url_status()

