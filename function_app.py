import json
import logging
from pathlib import Path

import azure.functions as func
import xlwings as xw

import custom_functions

app = func.FunctionApp()


@app.function_name(name="taskpane")
@app.route(route="taskpane.html", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def taskpane(req: func.HttpRequest):
    """This endpoint delivers the task pane incl. all client side HTML/JS/CSS"""
    taskpane = Path(__file__).parent / "taskpane.html"
    with open(taskpane, "r") as f:
        html_content = f.read()
    return func.HttpResponse(html_content, mimetype="text/html")


@app.function_name(name="custom-functions-meta")
@app.route(
    route="xlwings/custom-functions-meta",
    methods=["GET"],
    auth_level=func.AuthLevel.ANONYMOUS,
)
def custom_functions_meta(req: func.HttpRequest):
    """This endpoint delivers function signatures and descriptions"""
    return func.HttpResponse(
        json.dumps(xw.pro.custom_functions_meta(custom_functions)),
        mimetype="application/json",
    )


@app.function_name(name="custom-functions-code")
@app.route(
    route="xlwings/custom-functions-code",
    methods=["GET"],
    auth_level=func.AuthLevel.ANONYMOUS,
)
def custom_functions_code(req: func.HttpRequest):
    """This endpoint delivers the Office.js function wrappers"""
    return func.HttpResponse(
        xw.pro.custom_functions_code(
            custom_functions, "/api/xlwings/custom-functions-call"
        ),
        mimetype="text/javascript",
    )


@app.function_name(name="custom-functions-call")
@app.route(
    route="xlwings/custom-functions-call",
    methods=["POST"],
    auth_level=func.AuthLevel.ANONYMOUS,
)
async def custom_functions_call(req: func.HttpRequest):
    """This endpoint makes the Python function calls and can be protected via auth"""
    logging.info("custom_functions_call called")
    data = req.get_json()
    rv = await xw.pro.custom_functions_call(data, custom_functions)
    return func.HttpResponse(
        json.dumps({"result": rv}),
        mimetype="application/json",
    )
