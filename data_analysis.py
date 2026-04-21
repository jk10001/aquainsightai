import io
import json
import os
import platform
import random
import shutil
import string
import subprocess
import threading
import time
import traceback
import webbrowser
from datetime import datetime, timedelta, UTC
from pathlib import Path

# from pprint import pprint
from re import DOTALL, compile, escape
from typing import Annotated, Any, Dict, List, Optional, TypedDict, Union

import agentops
import autogen
import autogen.token_count_utils
import docker

import fitz  # PyMuPDF
import jinja2
import pandas as pd
import yaml
from autogen import Agent, ConversableAgent, register_function
from autogen.agentchat.contrib.multimodal_conversable_agent import (
    MultimodalConversableAgent,
)
from autogen.agentchat.contrib.society_of_mind_agent import SocietyOfMindAgent
from custom_multimodal_conversable_agent import CustomMultimodalConversableAgent
from autogen.coding import DockerCommandLineCodeExecutor
from docker.errors import DockerException, NotFound, APIError
from dotenv import load_dotenv
from flask import (
    Flask,
    jsonify,
    redirect,
    render_template,
    request,
    send_from_directory,
)
from flask_socketio import SocketIO
from termcolor import colored, cprint
from werkzeug.utils import secure_filename
from PIL import Image

from llm_cfg import (
    API_KEY_STATUS,
    ALIASES,
    ALIASES_VISION,
    MODEL_INFO,
    cfg,
)
from utils.lo_orchestrator import (
    process_docx_via_libreoffice,
    stop_lo_container,
)
from utils.scrub_docx_metadata import scrub_docx_metadata
from utils.scrub_pdf_metadata import scrub_pdf_metadata
from utils.csv_rounding import (
    EXCLUDE_KEYWORDS,
    OVERRIDE_INCLUDE_KEYWORDS,
    csv_sigfig_to_string,
)


load_dotenv()

# Flask setup
app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "data_store"
app.config["ALLOWED_EXTENSIONS"] = {"csv", "xlsx"}
socketio = SocketIO(app)

# Initial global variables
active_analysis = []
user_instructions = ""
user_response = None
user_additional_instructions = ""
response_event = threading.Event()
data_store: List[Dict[str, any]] = []
chart_or_table_in_production = ""
reference_docs = []
llm_report_review = False

# File names
HTML_INTERFACE_FILENAME = "data_analysis.html"
PROMPT_FILENAME = "data_analysis.yaml"
DATA_STORE_JSON_FILENAME = "data_store.json"
REFERENCE_DOCS_JSON_FILENAME = "reference_docs.json"

# Constants
LIBREOFFICE_DOCKER_IMAGE = "lo-headless:latest"
DATA_ANALYSIS_DOCKER_IMAGE = "data-analysis:latest"
DATA_DESCRIPTION_LARGE_TOKEN_LIMIT = 10000
DATA_INSERT_TOKEN_LIMIT = 10000
DATA_REMOVE_TOKEN_LIMIT = 2000
CHART_DATA_REPORT_TOKEN_LIMIT = 500
CHART_AND_TABLE_CODE_TOKEN_LIMIT = 4000
SILENT = False
AGENT_CACHE = None
RUN_AGENTOPS = False
TABLE_SIG_FIGS = 6
DEFAULT_USER_NOTES = "No notes provided."
ANTHROPIC_THINKING_ENABLE = True
ALLOW_HTML_SUMMARY_GENERATION = False
REMOVE_PREVIOUS_IMAGES = False
if ALLOW_HTML_SUMMARY_GENERATION:
    from utils.sanitise_html import sanitise_html


# Default LLMs
# process_agent - alias from llm_cfg ALIASES_VISION
DEFAULT_DATA_LLM_ALIAS = "GPT-5.2"
# int_data_coding_agent - alias from llm_cfg ALIASES
DEFAULT_CODE_LLM_ALIAS = "GPT-5 mini"
# summary_agent - alias from llm_cfg ALIASES_VISION
DEFAULT_SUMMARY_LLM_ALIAS = "Gemini 3 Flash"
# report_agent - alias from llm_cfg ALIASES_VISION
DEFAULT_REPORT_LLM_ALIAS = "GPT-5.2"
# Model for reference document query - must be a Gemini model name for reference document queries. This does not use alias pool - provide actual model name - "gemini-2.5-flash-lite" or "gemini-3-flash-preview"
QUERY_REFERENCE_DOC_LLM_NAME = "gemini-2.5-flash-lite"


def _resolve_default_alias(default_alias: str, alias_pool: list[str]) -> str:
    """Return a usable default alias, preferring one with a valid API key."""

    if API_KEY_STATUS.get(default_alias):
        return default_alias

    for alias in alias_pool:
        if API_KEY_STATUS.get(alias):
            cprint(
                f"API key missing for default alias '{default_alias}'. Using '{alias}' instead.",
                "cyan",
            )
            return alias

    cprint(
        f"No available API keys found for aliases: {', '.join(alias_pool)}."
        " The configured default '{default_alias}' remains unavailable.",
        "red",
    )
    return default_alias


DEFAULT_REPORT_LLM_ALIAS = _resolve_default_alias(
    DEFAULT_REPORT_LLM_ALIAS, ALIASES_VISION
)
DEFAULT_DATA_LLM_ALIAS = _resolve_default_alias(DEFAULT_DATA_LLM_ALIAS, ALIASES_VISION)
DEFAULT_CODE_LLM_ALIAS = _resolve_default_alias(DEFAULT_CODE_LLM_ALIAS, ALIASES)
DEFAULT_SUMMARY_LLM_ALIAS = _resolve_default_alias(
    DEFAULT_SUMMARY_LLM_ALIAS, ALIASES_VISION
)

if os.environ.get("GEMINI_API_KEY"):
    QUERY_REFERENCE_DOCS_ENABLED = True
else:
    QUERY_REFERENCE_DOCS_ENABLED = False
    cprint(
        "Warning: No Gemini API key found - reference document querying disabled.",
        "red",
    )


# Determine method for docx processing
def _libreoffice_image_available(image_name: str = LIBREOFFICE_DOCKER_IMAGE) -> bool:
    try:
        client = docker.from_env()
        client.images.get(image_name)
        return True
    except NotFound:
        cprint(
            f"LibreOffice Docker image '{image_name}' not found. Build or pull it before generating DOCX reports.",
            "red",
        )
    except DockerException as exc:
        cprint(
            f"Unable to access Docker while checking for LibreOffice image: {exc}",
            "red",
        )
    return False


def _select_libreoffice_processing() -> Optional[str]:
    if _libreoffice_image_available():
        cprint("Using LibreOffice for DOCX processing.", "cyan")
        return "libreoffice"

    cprint(
        "LibreOffice DOCX processing is disabled.",
        "red",
    )
    return None


process_docx_method: Optional[str] = None

if platform.system() == "Windows":
    try:
        import pythoncom
        import pywintypes
        import win32com.client as win32

        pythoncom.CoInitialize()
        word = None
        started = False
        try:
            # Try to bind to an existing Word instance
            try:
                word = win32.GetActiveObject("Word.Application")
                started = False
            except pywintypes.com_error:
                # No running instance
                word = win32.DispatchEx("Word.Application")
                started = True
                word.Visible = False
            process_docx_method = "word"
            cprint("Using Microsoft Word for DOCX processing", "cyan")

        finally:
            if started and word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    except Exception:
        process_docx_method = _select_libreoffice_processing()
else:
    process_docx_method = _select_libreoffice_processing()

if process_docx_method is None:
    cprint(
        "Warning: No method available for DOCX processing. LLM review of reports and conversion to PDF will be disabled.",
        "red",
    )


class sheets(TypedDict):
    sheet_name: str
    headings: list[str]
    time_start: str
    time_end: str
    time_step: Union[int, str]
    time_format: str


class PromptManager:
    def __init__(self, yaml_file: str):
        """
        Initialize PromptManager with a single YAML file.

        Args:
            yaml_file: Path to the YAML file containing prompts
        """
        self.yaml_file = Path(yaml_file)
        self.prompts: Dict[str, Any] = {}
        self.template_env = jinja2.Environment(
            loader=jinja2.FileSystemLoader(str(self.yaml_file.parent)),
            trim_blocks=True,
            lstrip_blocks=True,
        )
        self._load_prompts()

    def _load_prompts(self) -> None:
        """Load prompts from the specified YAML file."""
        if not self.yaml_file.exists():
            raise FileNotFoundError(f"Prompts file {self.yaml_file} not found")

        try:
            with open(self.yaml_file, "r", encoding="utf-8") as f:
                self.prompts = yaml.safe_load(f)
                cprint("Prompts loaded successfully.", "cyan")
        except UnicodeDecodeError as e:
            cprint(
                f"Warning: Encoding error in {
                    self.yaml_file
                }. Try saving the file in UTF-8 format. Error: {e}",
                "red",
            )
        except yaml.YAMLError as e:
            cprint(f"Warning: YAML parsing error in {self.yaml_file}: {e}", "red")
        except Exception as e:
            cprint(f"Warning: Unexpected error reading {self.yaml_file}: {e}", "red")

    def get_prompt(self, prompt_path: str, **kwargs) -> str:
        """
        Get a prompt by path and format it with the provided variables.

        Args:
            prompt_path: Path to the prompt using dot notation (e.g., 'chat.greeting.formal')
            **kwargs: Variables to format the prompt with

        Returns:
            Formatted prompt string

        Raises:
            KeyError: If the prompt path is invalid
            ValueError: If the prompt is not a string template
        """

        parts = prompt_path.split(".")

        # Navigate through the nested dictionary
        current = self.prompts
        for part in parts:
            if not isinstance(current, dict):
                raise KeyError(
                    f"Cannot navigate through non-dictionary value at '{part}'"
                )
            if part not in current:
                raise KeyError(f"Prompt path component '{part}' not found")
            current = current[part]

        if not isinstance(current, str):
            raise ValueError(f"The prompt at '{prompt_path}' is not a string template")

        template = self.template_env.from_string(current)
        return template.render(**kwargs)


def load_data_store():
    """Load data from JSON file."""
    global data_store
    try:
        with open(
            os.path.join(app.config["UPLOAD_FOLDER"], DATA_STORE_JSON_FILENAME),
            "r",
            encoding="utf-8",
        ) as f:
            data_store = json.load(f)
            for analysis in data_store:
                analysis["processing"] = False
                analysis.setdefault("user_notes", DEFAULT_USER_NOTES)
        cprint("Data store loaded successfully.", "cyan")
    except FileNotFoundError:
        data_store = []
        cprint("No existing data store found - starting new.", "red")


def save_data_store():
    """Save data to JSON file."""
    with open(
        os.path.join(app.config["UPLOAD_FOLDER"], DATA_STORE_JSON_FILENAME),
        "w",
        encoding="utf-8",
    ) as f:
        json.dump(data_store, f, ensure_ascii=False)


def load_reference_docs():
    """Load data from JSON file."""
    global reference_docs
    try:
        with open(
            os.path.join("reference_docs", REFERENCE_DOCS_JSON_FILENAME), "r"
        ) as f:
            reference_docs = json.load(f)
        cprint("Reference documents loaded successfully.", "cyan")
    except FileNotFoundError:
        reference_docs = []


def save_reference_docs():
    """Save data to JSON file."""
    global reference_docs
    with open(os.path.join("reference_docs", REFERENCE_DOCS_JSON_FILENAME), "w") as f:
        json.dump(reference_docs, f)
    cprint("Reference documents saved successfully.", "yellow")


def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in app.config["ALLOWED_EXTENSIONS"]
    )


def add_data_description(
    data_description: Annotated[str, "Brief description of the data file."],
    sheet_list: List[sheets],
    analysis_name: Annotated[
        str,
        "A concise title for the analysis (approximately 35 characters or fewer).",
    ],
) -> str:
    """Function called by process_agent to record the initial description of the provided data file."""
    global user_instructions, user_response, user_additional_instructions
    current_analysis = data_store[-1]
    description_payload = {
        "data_description": data_description,
        "sheets": sheet_list,
    }
    current_analysis["description"] = description_payload
    current_analysis["name"] = analysis_name
    save_data_store()
    socketio.emit(
        "data_description_update",
        {"description": description_payload, "analysis_name": analysis_name},
    )

    # Update process_agent system message and tools
    process_agent.update_system_message(
        prompt_manager.get_prompt(
            "default.process_agent_system_message_analysis",
            user_instructions=user_instructions,
        )
    )
    process_agent.update_tool_signature("add_data_description", is_remove=True)
    register_function(
        add_chart_and_comment,
        caller=process_agent,
        executor=function_agent,
        name="add_chart_and_comment",
        description="A function to record an analysis step when the output is acceptable, including a title, analysis comments, chart image filename (optional), chart html filename (optional), code filename, data filename and txt filename with additional information (optional). If a filename parameter is not provided, leave the argument out. Do not ask for changes to the chart or table via this function call.",
    )

    query_reference_doc_description = prompt_manager.get_prompt(
        "default.tools.reference_doc_tool_description",
        reference_docs=reference_docs,
    )

    if QUERY_REFERENCE_DOCS_ENABLED:
        register_function(
            query_reference_doc,
            caller=process_agent,
            executor=function_agent,
            name="query_reference_doc",
            description=query_reference_doc_description,
        )

    socketio.emit(
        "show_popup",
        {
            "line1": "Initial data review complete. Proceed with data analysis?",
            "line2": "",
            "buttons": ["Continue", "Stop and summarise"],
            "callback_function": "handle_data_description_response",
        },
    )
    global chart_or_table_in_production
    chart_or_table_in_production = ""

    response = wait_for_user_response()
    additional_instructions = user_additional_instructions.strip()
    additional_instructions_message = (
        f"\nUser's additional instructions for next step: {additional_instructions}"
        if additional_instructions
        else ""
    )

    if response == "Continue":
        return (
            "Result from tool call: Successfully added data descriptions. "
            f"Now proceed with requesting charts or tables. {additional_instructions_message}"
        )
    else:
        return "TERMINATE"


def add_chart_and_comment(
    title: Annotated[str, "Title of the chart or table."],
    analysis_comments: Annotated[
        str,
        "Analysis comments for the chart or table. Do not mention names of files created.",
    ],
    chart_filename: Annotated[
        Optional[str], "Filename of the chart PNG, if created."
    ] = None,
    chart_html_filename: Annotated[
        Optional[str], "Filename of the chart HTML, if created."
    ] = None,
    code_filename: Annotated[
        Optional[str], "Filename of the python file used in the analysis."
    ] = None,
    data_filename: Annotated[
        Optional[str], "Filename of the CSV file created in the analysis."
    ] = None,
    additional_info_filename: Annotated[
        Optional[str],
        "Filename of the TXT file giving additional information, if created.",
    ] = None,
) -> str:
    """
    Function called by process_agent to record a chart or table and comments.
    Called during the main analysis routine.
    """
    global user_response, user_additional_instructions
    current_analysis = data_store[-1]
    new_entry = {
        "title": title,
        "chart_filename": chart_filename,
        "chart_html_filename": chart_html_filename,
        "analysis_comments": analysis_comments,
        "code_filename": code_filename,
        "data_filename": data_filename,
        "additional_info_filename": additional_info_filename,
    }

    current_analysis["charts_and_comments"].append(new_entry)
    current_analysis["summary"]["update_needed"] = True
    save_data_store()
    socketio.emit("new_chart", new_entry)
    global chart_or_table_in_production
    chart_or_table_in_production = ""

    socketio.emit(
        "show_popup",
        {
            "line1": "Continue with data analysis?",
            "line2": "",
            "buttons": ["Continue", "Stop and summarise"],
            "callback_function": "handle_chart_response",
        },
    )

    response = wait_for_user_response()
    additional_instructions = user_additional_instructions.strip()
    additional_instructions_message = (
        f"\nUser's additional instructions for next step: {additional_instructions}"
        if additional_instructions
        else ""
    )

    if response == "Continue":
        return (
            f"Result from tool call: Successfully added chart '{title}' and comments to list. "
            f"Now select another chart or table to produce. {additional_instructions_message}"
        )
    else:
        return "TERMINATE"


def revise_chart(
    analysis_index: int,
    item_index: int,
    title: Annotated[str, "Title of the chart or table."],
    analysis_comments: Annotated[
        str,
        "Analysis comments for the chart or table. Do not discuss how the chart/table has been revised from what it was previously, focus on analysis. Do not mention names of files created.",
    ],
    chart_filename: Annotated[
        Optional[str], "Filename of the chart PNG, if created."
    ] = None,
    chart_html_filename: Annotated[
        Optional[str], "Filename of the chart HTML, if created."
    ] = None,
    code_filename: Annotated[
        Optional[str], "Filename of the python file used in the analysis."
    ] = None,
    data_filename: Annotated[
        Optional[str], "Filename of the CSV file created in the analysis."
    ] = None,
    additional_info_filename: Annotated[
        Optional[str],
        "Filename of the MD file giving additional information, if created.",
    ] = None,
) -> str:
    """Function called by process_agent to record a revised chart or table and comments."""
    current_chart = data_store[analysis_index]["charts_and_comments"][item_index]
    current_chart["title"] = title
    current_chart["chart_filename"] = chart_filename
    current_chart["chart_html_filename"] = chart_html_filename
    current_chart["analysis_comments"] = analysis_comments
    current_chart["code_filename"] = code_filename
    current_chart["data_filename"] = data_filename
    current_chart["additional_info_filename"] = additional_info_filename

    global chart_or_table_in_production
    chart_or_table_in_production = ""
    data_store[analysis_index].setdefault("report", {})["update_needed"] = True
    data_store[analysis_index]["summary"]["update_needed"] = True
    save_data_store()

    return "TERMINATE"


def add_new_chart(
    analysis_index: int,
    title: Annotated[str, "Title of the chart or table."],
    analysis_comments: Annotated[
        str,
        "Analysis comments for the chart or table. Do not mention names of files created.",
    ],
    chart_filename: Annotated[
        Optional[str], "Filename of the chart PNG, if created."
    ] = None,
    chart_html_filename: Annotated[
        Optional[str], "Filename of the chart HTML, if created."
    ] = None,
    code_filename: Annotated[
        Optional[str], "Filename of the python file used in the analysis."
    ] = None,
    data_filename: Annotated[
        Optional[str], "Filename of the CSV file created in the analysis."
    ] = None,
    additional_info_filename: Annotated[
        Optional[str],
        "Filename of the MD file giving additional information, if created.",
    ] = None,
) -> str:
    """Function called by process_agent to record a chart or table and comments.
    Called during the add new chart routine."""
    current_analysis = data_store[analysis_index]
    new_entry = {
        "title": title,
        "chart_filename": chart_filename,
        "chart_html_filename": chart_html_filename,
        "analysis_comments": analysis_comments,
        "code_filename": code_filename,
        "data_filename": data_filename,
        "additional_info_filename": additional_info_filename,
    }
    current_analysis["charts_and_comments"].append(new_entry)
    current_analysis.setdefault("report", {})["update_needed"] = True
    current_analysis["summary"]["update_needed"] = True
    save_data_store()

    socketio.emit(
        "new_chart_added",
        {
            "analysisIndex": analysis_index,
            "newChart": new_entry,
            "update_needed": True,
        },
    )
    global chart_or_table_in_production
    chart_or_table_in_production = ""

    return "TERMINATE"


def wait_for_user_response():
    """Wait for user response in popup."""
    global user_response, user_additional_instructions, response_event
    user_response = None
    user_additional_instructions = ""
    response_event.clear()
    response_event.wait()  # This will block until the event is set
    return user_response


@socketio.on("user_response")
def handle_user_response(data):
    """Called from html when user clicks on popup."""
    global user_response, user_additional_instructions, response_event
    user_response = data["response"]
    user_additional_instructions = data.get("additional_instructions", "")
    response_event.set()  # This will unblock the waiting thread


def _format_chart_or_table_context(
    *,
    item: dict[str, Any],
    analysis: dict[str, Any],
    include_images: bool,
    include_code_token_limit: Optional[int],
    include_chart_csv_token_limit: Optional[int],
    include_csv_filename: bool,
    include_chart_image_filename: bool,
) -> str:
    """Format a single chart or table into markdown context."""

    lines: list[str] = []
    component_label = "Chart" if item["chart_filename"] is not None else "Table"
    lines.append(f"####{component_label}: {item['title']}")
    lines.append(item["analysis_comments"].strip())

    if include_chart_image_filename and item.get("chart_filename"):
        lines.append(f"Chart image file name: {item['chart_filename']}")
    if include_images and item.get("chart_filename"):
        lines.append(
            f"Chart image: <img {analysis['data_folder']}/{item['chart_filename']}>."
        )

    if (
        item.get("chart_filename") is not None
        and include_chart_csv_token_limit is not None
    ):
        if item.get("data_filename") is not None:
            chart_data = csv_sigfig_to_string(
                os.path.join(analysis["data_folder"], item["data_filename"]),
                sig_figs=TABLE_SIG_FIGS,
            )
            if (
                autogen.token_count_utils.count_token(chart_data, model="gpt-4o")
                <= include_chart_csv_token_limit
            ):
                if include_csv_filename:
                    lines.append(f"Chart data file name: {item['data_filename']}")
                lines.append(f"Chart data table:\n{chart_data}")

    if item.get("chart_filename") is None:
        if include_csv_filename and item.get("data_filename"):
            lines.append(f"Table data file name: {item['data_filename']}")
        if item.get("data_filename") is not None:
            table_data = csv_sigfig_to_string(
                os.path.join(analysis["data_folder"], item["data_filename"]),
                sig_figs=TABLE_SIG_FIGS,
            )
            lines.append(table_data)

    if include_code_token_limit is not None and item.get("code_filename"):
        code_path = os.path.join(analysis["data_folder"], item["code_filename"])
        code = load_file_as_string(code_path)
        if (
            autogen.token_count_utils.count_token(code, model="gpt-4o")
            <= include_code_token_limit
        ):
            lines.append(f"Code used to create component:\n```python\n{code}\n```")

    return "\n".join([line for line in lines if line]).strip()


def chart_and_table_context_builder(
    analysis_index: int,
    include_images: bool,
    include_code_token_limit: Optional[int],
    include_chart_csv_token_limit: Optional[int],
    include_csv_filename: bool,
    include_chart_image_filename: bool,
    exclude_component_index: Optional[int] = None,
) -> str:
    """Build ordered markdown context for charts and tables in an analysis."""

    analysis = data_store[analysis_index]
    context_sections: list[str] = []

    for idx, item in enumerate(analysis["charts_and_comments"]):
        if exclude_component_index is not None and idx == exclude_component_index:
            continue

        context_sections.append(
            _format_chart_or_table_context(
                item=item,
                analysis=analysis,
                include_images=include_images,
                include_code_token_limit=include_code_token_limit,
                include_chart_csv_token_limit=include_chart_csv_token_limit,
                include_csv_filename=include_csv_filename,
                include_chart_image_filename=include_chart_image_filename,
            )
        )

    return "\n\n".join(section for section in context_sections if section)


def custom_speaker_selection_func(last_speaker: Agent, groupchat: autogen.GroupChat):
    """Customized speaker selection function for the main groupchat.
    Parameters:
        - last_speaker: Agent - the last speaker in the group chat.
        - groupchat: GroupChat - the GroupChat object
    Return:
        The next agent
    """
    messages = groupchat.messages
    if last_speaker is process_agent and "tool_calls" in messages[-1].keys():
        if messages[-1]["tool_calls"] is not None:
            return function_agent
        else:
            return code_som_agent
    elif last_speaker is code_som_agent:
        return process_agent
    elif last_speaker is function_agent:
        return process_agent
    return code_som_agent


def hook_remove_old_data_from_message(all_messages):
    """Hook function to edit incoming process_agent message before it is sent to the LLM (process_all_messages_before_reply).
    For all messages except the latest, data previously added inside <dat   > will be removed if above the token limit.
    """
    # Regular expression pattern for matching <dat ...> tags
    dat_tag_pattern_2 = compile(r"<dat (.*?)>", DOTALL)
    for message in all_messages[:-1]:
        if "content" in message and message["content"] is not None:
            if isinstance(message["content"], list):
                for content_item in message["content"]:
                    if "text" in content_item:
                        # Remove the pattern from the text value
                        dat_match = dat_tag_pattern_2.search(content_item["text"])
                        if dat_match:
                            data_tokens = autogen.token_count_utils.count_token(
                                input=dat_match.group(1), model="gpt-4"
                            )
                            if data_tokens > DATA_REMOVE_TOKEN_LIMIT:
                                content_item["text"] = dat_tag_pattern_2.sub(
                                    "", content_item["text"]
                                )

            elif isinstance(message["content"], str):
                # Remove the pattern from the text value
                dat_match = dat_tag_pattern_2.search(message["content"])
                if dat_match:
                    data_tokens = autogen.token_count_utils.count_token(
                        input=dat_match.group(1), model="gpt-4"
                    )
                    if data_tokens > DATA_REMOVE_TOKEN_LIMIT:
                        message["content"] = dat_tag_pattern_2.sub(
                            "", message["content"]
                        )
    return all_messages


def docker_start(folder) -> tuple[Optional[DockerCommandLineCodeExecutor], bool]:
    """
    Start a docker container in the provided folder and set up executor.
    Returns: tuple of executor and error flag.
    """
    cprint("Checking Docker...", "cyan")
    client = None
    while True:
        try:
            client = docker.from_env()
            client.ping()
            cprint("Docker engine is running.", "cyan")
            break
        except DockerException as e:
            cprint(
                f"Docker engine is not running. Please start Docker and try again.\nError: {e}",
                "red",
            )
            input("Start Docker then press return: ")
            time.sleep(1)

    container_name = "data-analysis"
    try:
        try:
            client.images.get(DATA_ANALYSIS_DOCKER_IMAGE)
        except NotFound:
            cprint(
                f"Docker image '{DATA_ANALYSIS_DOCKER_IMAGE}' was not found. "
                "Please build the required image and restart.",
                "red",
            )
            return None, True
        except DockerException as e:
            cprint(
                f"Unable to access Docker images.\nError: {e}",
                "red",
            )
            return None, True

        while True:
            try:
                container = client.containers.get(container_name)
            except NotFound:
                cprint(
                    f"Container '{container_name}' does not exist. Container will be created.",
                    "cyan",
                )
                break

            cprint(f"Removing container '{container_name}'...", "cyan")
            try:
                container.remove(force=True)
            except APIError as e:
                cprint(
                    f"Warning: remove() raised {e}. Will still poll for deletion...",
                    "red",
                )

            cprint("Waiting for container to be deleted...", "cyan")
            for attempt in range(1, 16):
                try:
                    client.containers.get(container_name)
                    time.sleep(1)
                except NotFound:
                    cprint(
                        f"Container '{container_name}' has been deleted. New container will be created.",
                        "cyan",
                    )
                    break
            else:
                # only happens if polling loop exhausted
                input(
                    colored(
                        "Container still present after 15 seconds.\n"
                        "Please delete it manually in Docker, then press Enter to try again...",
                        "red",
                    )
                )
                continue

            break
    finally:
        if client is not None:
            client.close()

    try:
        executor = DockerCommandLineCodeExecutor(
            container_name=container_name,
            timeout=300,
            work_dir=folder,
            image=DATA_ANALYSIS_DOCKER_IMAGE,
            auto_remove=True,
            stop_container=True,
        )
    except Exception as e:
        cprint(f"Error starting Docker container: {e}", "red")
        print(traceback.format_exc())
        return None, True

    return executor, False


def load_file_as_string(file_path: str) -> str:
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            content = file.read()
        return content
    except Exception as e:
        cprint(f"Warning! Unable to load file {file_path}, because of {e}", "red")
        return f"Warning! Unable to load file {file_path}, because of {e}"


def get_next_filename(folder, prefix="coding_session_history_", suffix=".json"):
    if not os.path.exists(folder):
        os.makedirs(folder, exist_ok=True)
    existing_files = os.listdir(folder)
    pattern = compile(rf"{escape(prefix)}(\d+){escape(suffix)}$")
    numbers = [
        int(match.group(1)) for f in existing_files if (match := pattern.match(f))
    ]
    next_num = max(numbers, default=0) + 1
    return os.path.join(folder, f"{prefix}{next_num:02d}{suffix}")


def code_som_agent_responder(agent, messages):
    """Load information from completed coding task and format for process agent."""

    filename = get_next_filename(
        os.path.join(active_analysis["data_folder"], "chat_history"),
        prefix="coding_session_history_",
        suffix=".json",
    )
    with open(
        filename,
        "w",
    ) as f:
        json.dump(sanitize_for_json(messages), f, indent=4)

    last_message_content = messages[-1]["content"]
    if "datadescription.md" in last_message_content:
        file_path = os.path.join(active_analysis["data_folder"], "datadescription.md")
        with open(file_path, "r") as file:
            file_contents = file.read()
            return f"<code_ouptut_result_agent>:\n {file_contents}"
    else:
        img_tag_pattern = compile(r"<img ([^>]+)>")
        htm_tag_pattern = compile(r"<htm ([^>]+)>")
        pyt_tag_pattern = compile(r"<pyt ([^>]+)>")
        dat_tag_pattern = compile(r"<dat ([^>]+)>")
        inf_tag_pattern = compile(r"<inf ([^>]+)>")

        # extract data filename from dat tags and load data file
        data_match = dat_tag_pattern.search(last_message_content)
        if data_match:
            data_filename = data_match.group(1)
            data_location = active_analysis["data_folder"] + "/" + data_match.group(1)
            data = csv_sigfig_to_string(data_location, sig_figs=TABLE_SIG_FIGS)
            data_tokens = autogen.token_count_utils.count_token(
                input=data, model="gpt-4"
            )

            if data_tokens < DATA_INSERT_TOKEN_LIMIT:
                pass
            else:
                print(
                    f"Data longer than {
                        DATA_INSERT_TOKEN_LIMIT
                    } tokens. Data will not be included in prompt."
                )
                data = "Data too large - not provided."

        else:
            data = data_filename = data_location = "No data tag found."

        # extract additional info filename from inf tags and add path
        inf_match = inf_tag_pattern.search(last_message_content)
        if inf_match:
            inf_filename = inf_match.group(1)
            inf_location = active_analysis["data_folder"] + "/" + inf_match.group(1)
            info = load_file_as_string(inf_location)
        else:
            info = inf_location = "No additional information provided."
            inf_filename = "not applicable"

        # extract image filename from img tags and add path
        img_match = img_tag_pattern.search(last_message_content)
        if img_match:
            image_filename = img_match.group(1)
            image_with_path = active_analysis["data_folder"] + "/" + image_filename
        else:
            image_filename = image_with_path = "No image tag found."

        # extract chart html filename from htm tags
        htm_match = htm_tag_pattern.search(last_message_content)
        if htm_match:
            html_filename = htm_match.group(1)
        else:
            html_filename = "No html tag found."

        # extract code filename from pyt tags and load code file
        pyt_match = pyt_tag_pattern.search(last_message_content)
        if pyt_match:
            code_filename = pyt_match.group(1)
            code_location = active_analysis["data_folder"] + "/" + code_filename
            code = load_file_as_string(code_location)
        else:
            code = code_filename = "No code tag found."

        if image_filename == "No image tag found.":
            response = f"""<code_output_result_agent>:
Here is the table you requested: 
<dat d
This table saved in {data_filename}):
{data}>
Additional information (stored in {inf_filename}):
{info}
This table was created from this code (stored in {code_filename}):
{code}

Let me know if you would like any changes or improvements. If the table is ok, make a tool call to add_chart_and_comment.'"""
        else:
            response = f"""<code_output_result_agent>:
Here is the chart you requested (chart filename is {image_filename}): <img {image_with_path}>.
The html version of this chart is available at {html_filename}.

<dat d
This chart was created from this data (stored in {data_filename}):
{data}>
Additional information (stored in {inf_filename}):
{info}
This chart was created from this code (stored in {code_filename}):
{code}

Let me know if you would like any changes or improvements. If the chart is ok, make a tool call to add_chart_and_comment.'"""
    return response


def add_agent_name_to_message(
    sender: ConversableAgent,
    message: Union[dict[str, Any], str],
    recipient: Agent,
    silent: bool,
) -> Union[dict[str, Any], str]:
    """Add agent name to the start of the message."""

    global chart_or_table_in_production
    if sender.name == "process_agent":
        if isinstance(message, str):
            title_tag_pattern = compile(r"<title ([^>]+)>")
            title_match = title_tag_pattern.search(message)
            if title_match:
                chart_or_table_in_production = title_match.group(1)
            if message.startswith("<process_agent>:"):
                return message
            message = f"<process_agent>:\n{message}"
            return message
        elif isinstance(message, dict):
            if "content" in message:
                if isinstance(message["content"], str):
                    title_tag_pattern = compile(r"<title ([^>]+)>")
                    title_match = title_tag_pattern.search(message["content"])
                    if title_match:
                        chart_or_table_in_production = title_match.group(1)
                    if message["content"].startswith("<process_agent>:"):
                        return message
                    message["content"] = f"<process_agent>:\n{message['content']}"
                    return message

    if sender.name == "int_data_coding_agent":
        if isinstance(message, str):
            if message.startswith("<coding_agent>:"):
                return message
            message = f"<coding_agent>:\n{message}"
            return message
        elif isinstance(message, dict):
            if message.get("content"):
                if isinstance(message["content"], str):
                    if message["content"].startswith("<coding_agent>:"):
                        return message
                    message["content"] = f"<coding_agent>:\n{message['content']}"
                    return message
    if sender.name == "function_agent":
        return message
    cprint("Error adding agent name to message.", "red")
    return message


def initialise_agents(
    purpose,
    active_folder,
    data_llm_alias: str = DEFAULT_DATA_LLM_ALIAS,
    code_llm_alias: str = DEFAULT_CODE_LLM_ALIAS,
    data_llm_params: Optional[Dict[str, Any]] = None,
    code_llm_params: Optional[Dict[str, Any]] = None,
) -> bool:
    """Initilise agents for main analysis process or revise/add chart.
    - int_data_coding_agent
    - int_code_interpreter
    - int_groupchat
    - int_manager
    - code_som_agent
    - process_agent
    - function_agent
    - main_groupchat
    - main_manager
    """
    global \
        process_agent, \
        main_manager, \
        int_data_coding_agent, \
        int_code_interpreter, \
        int_groupchat, \
        int_manager, \
        code_som_agent, \
        function_agent, \
        main_groupchat, \
        main_groupchat, \
        executor, \
        reference_docs

    try:
        executor = None

        message = "Initialising..."
        cprint(message, "cyan", attrs=["bold"])
        socketio.emit("update_popup", {"line1": message})

        cprint(
            f"Data analysis LLM:\n{cfg(data_llm_alias, **(data_llm_params or {}))}\n",
            "cyan",
        )
        cprint(
            f"Code analysis LLM:\n{cfg(code_llm_alias, **(code_llm_params or {}))}\n",
            "cyan",
        )

        if purpose == "mainprocess":
            process_agent_system_message = prompt_manager.get_prompt(
                "default.process_agent_system_message_data_investigation"
            )
            process_agent_max_consecutive_auto_reply = 12
            main_groupchat_max_round = 20
        elif purpose == "revisechart" or purpose == "addnewchart":
            process_agent_system_message = "xxxxxx"
            process_agent_max_consecutive_auto_reply = 8
            main_groupchat_max_round = 12
        else:
            cprint("Error initialising agents.\n\n", "red", attrs=["bold"])
            socketio.emit(
                "show_popup",
                {
                    "line1": "Error initialising, terminating.",
                    "buttons": ["Ok"],
                },
            )
            wait_for_user_response()
            return False

        executor, docker_error = docker_start(active_folder)
        if docker_error:
            cprint("Docker not available, terminating.\n\n", "red", attrs=["bold"])
            socketio.emit(
                "show_popup",
                {
                    "line1": "Docker not available, terminating.",
                    "buttons": ["Ok"],
                },
            )
            wait_for_user_response()
            executor = None
            return False

        process_agent = CustomMultimodalConversableAgent(
            name="process_agent",
            max_consecutive_auto_reply=process_agent_max_consecutive_auto_reply,
            system_message=process_agent_system_message,
            code_execution_config=False,
            remove_previous_images=REMOVE_PREVIOUS_IMAGES,
            llm_config={
                "cache_seed": AGENT_CACHE,
                "config_list": cfg(data_llm_alias, **(data_llm_params or {})),
                # "temperature": 0.3,
                # "max_tokens": 4000,
            },
            human_input_mode="NEVER",
        )

        function_agent = autogen.UserProxyAgent(
            "function_agent",
            human_input_mode="NEVER",
            code_execution_config=False,
            llm_config=False,
            default_auto_reply="",
        )

        int_data_coding_agent = autogen.AssistantAgent(
            name="int_data_coding_agent",
            system_message=prompt_manager.get_prompt(
                "default.int_coding_agent_system_message_plotly"
            ),
            max_consecutive_auto_reply=8,
            llm_config={
                "cache_seed": AGENT_CACHE,
                "config_list": cfg(code_llm_alias, **(code_llm_params or {})),
                # "temperature": 0.1,
            },
        )

        int_code_interpreter = autogen.UserProxyAgent(
            "int_code_interpreter",
            human_input_mode="NEVER",
            code_execution_config={"executor": executor},
            llm_config=False,
            default_auto_reply="",
        )

        int_groupchat = autogen.GroupChat(
            agents=[int_data_coding_agent, int_code_interpreter],
            messages=[],
            speaker_selection_method="round_robin",
            allow_repeat_speaker=False,
            max_round=12,
        )

        int_manager = autogen.GroupChatManager(
            groupchat=int_groupchat,
            is_termination_msg=lambda msg: "TASK COMPLETE" in msg["content"].upper(),
            llm_config=False,
        )

        code_som_agent = SocietyOfMindAgent(
            "code_som_agent",
            chat_manager=int_manager,
            response_preparer=code_som_agent_responder,
            llm_config=False,
        )

        main_groupchat = autogen.GroupChat(
            agents=[process_agent, code_som_agent, function_agent],
            messages=[],
            speaker_selection_method=custom_speaker_selection_func,
            allow_repeat_speaker=False,
            max_round=main_groupchat_max_round,
        )
        main_manager = autogen.GroupChatManager(
            groupchat=main_groupchat,
            llm_config=False,
            is_termination_msg=lambda msg: (
                "terminate"
                in (
                    msg["content"][0]["text"].lower()
                    if isinstance(msg.get("content"), list)
                    and len(msg["content"]) > 0
                    and isinstance(msg["content"][0], dict)
                    and "text" in msg["content"][0]
                    else msg["content"].lower()
                    if isinstance(msg.get("content"), str)
                    else ""
                )
            ),
        )

        # Register hooks for processing messages.
        # Available hookable methods: "process_last_received_message", "process_all_messages_before_reply" and "process_message_before_send"
        process_agent.register_hook(
            hookable_method="process_all_messages_before_reply",
            hook=hook_remove_old_data_from_message,
        )
        int_data_coding_agent.register_hook(
            hookable_method="process_all_messages_before_reply",
            hook=hook_remove_old_data_from_message,
        )

        # Register hooks for updating status on webpage
        process_agent.register_hook(
            hookable_method="process_all_messages_before_reply",
            hook=update_process_status_process_agent,
        )
        int_data_coding_agent.register_hook(
            hookable_method="process_all_messages_before_reply",
            hook=update_process_status_int_data_coding_agent,
        )
        int_code_interpreter.register_hook(
            hookable_method="process_all_messages_before_reply",
            hook=update_process_status_int_code_interpreter,
        )

        # Register hooks for adding agent name to message
        # code_som_agent also has name added as part of code_som_agent_responder()
        process_agent.register_hook(
            hookable_method="process_message_before_send",
            hook=add_agent_name_to_message,
        )
        int_data_coding_agent.register_hook(
            hookable_method="process_message_before_send",
            hook=add_agent_name_to_message,
        )

        # Register functions
        if purpose == "mainprocess":
            register_function(
                add_data_description,
                caller=process_agent,
                executor=function_agent,
                name="add_data_description",
                description="A function to record information about the data file, including a description and, for each sheet in the file: the 'sheet_name', list of headings ('headings'), date and time when the data starts ('time_start'), date and time when the data stops ('time_end'), the 'time_step' between datapoints and the date and time format ('time_format'). If time related arguments only if applicable, provide 'Not applicable' as the argument. Also include an 'analysis_name' argument that provides a concise title of roughly 35 characters or fewer.",
            )
        elif purpose == "revisechart":
            register_function(
                revise_chart,
                caller=process_agent,
                executor=function_agent,
                name="revise_chart",
                description="A function to update a revised chart or table, commentary and the filenames of the data, code and and additional information for the chart when the chart or table is acceptable. If a filename is not provided, leave the argument out.",
            )
        elif purpose == "addnewchart":
            register_function(
                add_new_chart,
                caller=process_agent,
                executor=function_agent,
                name="add_new_chart",
                description="A function to record a new chart or table, commentary and the filenames of the  data, code and and additional information for the chart when the chart or table is acceptable. If a filename is not provided, leave the argument out.",
            )
        else:
            cprint("Error initialising agents.\n\n", "red", attrs=["bold"])
            socketio.emit(
                "show_popup",
                {
                    "line1": "Error initialising, terminating.",
                    "buttons": ["Ok"],
                },
            )
            wait_for_user_response()
            return False

        if purpose == "revisechart" or purpose == "addnewchart":
            query_reference_doc_description = prompt_manager.get_prompt(
                "default.tools.reference_doc_tool_description",
                reference_docs=reference_docs,
            )
            if QUERY_REFERENCE_DOCS_ENABLED:
                register_function(
                    query_reference_doc,
                    caller=process_agent,
                    executor=function_agent,
                    name="query_reference_doc",
                    description=query_reference_doc_description,
                )
            else:
                cprint(
                    "Gemini API key not configured. Reference document queries disabled.",
                    "yellow",
                )

        cprint("Agents initialised.\n\n", "cyan")
        return True

    except Exception as e:
        cprint(f"Error initialising agents: {e}\n\n", "red", attrs=["bold"])
        print(traceback.format_exc())
        socketio.emit(
            "show_popup",
            {
                "line1": "Error initialising, terminating.",
                "buttons": ["Ok"],
            },
        )
        wait_for_user_response()
        return False


def query_reference_doc(
    doc_index: Annotated[int, "Index number for document."],
    prompt: Annotated[str, "LLM prompt to retrieve information from document."],
) -> str:
    """
    Function to call the reference document agent.
    Args:
        prompt: str - the prompt to send to the agent
        doc_index: int - the index of the document to query
    Returns:
        str: the response from the agent
    """
    message = f"Querying document: {reference_docs[doc_index]['doc_name']}..."
    cprint(message, "yellow", attrs=["bold"])
    socketio.emit(
        "update_popup", {"line1": message, "line2": "This may take a while..."}
    )

    from google import genai

    gemini_api_key = os.environ.get("GEMINI_API_KEY")

    client = genai.Client(api_key=gemini_api_key)

    document = upload_reference_doc(
        client=client, doc_index=doc_index, model_name=QUERY_REFERENCE_DOC_LLM_NAME
    )
    if document is None:
        cprint(
            f"Failed to upload or retrieve document {reference_docs[doc_index]['file_name']}.",
            "red",
        )
        return "Error: Document upload failed."

    cprint("Request underway\n", "yellow")
    response = client.models.generate_content(
        model=QUERY_REFERENCE_DOC_LLM_NAME,
        contents=[
            "Provided here is a reference document. Answer this request from the user: "
            + prompt,
            document,
        ],
    )
    cprint("Response recieved.", "yellow")
    return response.text


def upload_reference_doc(client, doc_index: int, model_name: str):
    """Check uploaded reference docs and re-upload if needed."""
    global reference_docs
    upload_date = reference_docs[doc_index].get("upload_date")
    if (
        upload_date is None
        or upload_date == ""
        or datetime.now(UTC) - datetime.fromisoformat(upload_date).replace(tzinfo=UTC)
        > timedelta(hours=48)
    ):
        path = os.path.join("reference_docs", reference_docs[doc_index]["file_name"])
        cprint(
            f"Re-uploading file {reference_docs[doc_index]['file_name']}...", "yellow"
        )
        try:
            obj = client.files.upload(file=path)
            tokens = client.models.count_tokens(model=model_name, contents=[obj])
            cprint(
                f"Document tokens: {tokens.total_tokens}",
                "yellow",
            )
            reference_docs[doc_index]["remote_name"] = obj.name
            reference_docs[doc_index]["upload_date"] = datetime.now(UTC).isoformat()
            cprint(f"Re-uploaded {obj.name}. Tokens: {tokens.total_tokens}.", "yellow")
            save_reference_docs()
            return obj
        except Exception as e:
            cprint(f"Failed to re-upload {path}: {e}", "red")
    else:
        try:
            file = client.files.get(name=reference_docs[doc_index]["remote_name"])
            cprint(
                f"Document {reference_docs[doc_index]['file_name']} already uploaded on {upload_date}.",
                "yellow",
            )
            return file
        except Exception as e:
            cprint(
                f"Failed to retrieve {reference_docs[doc_index]['remote_name']}: {e}",
                "red",
            )
            return None


def update_process_status_process_agent(all_messages):
    global chart_or_table_in_production
    if chart_or_table_in_production == "":
        message = """Analysing results and deciding next steps..."""
    else:
        message = f"""Producing {chart_or_table_in_production}\nAnalysing results and deciding next steps..."""
    cprint(message, "green", attrs=["bold"])
    socketio.emit("update_popup", {"message": message})
    return all_messages


def update_process_status_int_data_coding_agent(all_messages):
    global chart_or_table_in_production
    if "content" in all_messages[-1] and "exitcode: 0" in all_messages[-1]["content"]:
        if chart_or_table_in_production == "":
            message = """Code executed successfully..."""
        else:
            message = f"""Producing {chart_or_table_in_production}\nCode executed successfully..."""
    else:
        if chart_or_table_in_production == "":
            message = """Generating code..."""
        else:
            message = (
                f"""Producing {chart_or_table_in_production}\nGenerating code..."""
            )
    cprint(message, "blue", attrs=["bold"])
    socketio.emit("update_popup", {"message": message})
    return all_messages


def update_process_status_int_code_interpreter(all_messages):
    global chart_or_table_in_production
    if chart_or_table_in_production == "":
        message = """Executing code..."""
    else:
        message = f"""Producing {chart_or_table_in_production}\nExecuting code..."""
    cprint(message, "magenta", attrs=["bold"])
    socketio.emit("update_popup", {"message": message})
    return all_messages


def revise_chart_process(
    analysis_index: int,
    item_index: int,
    instructions: str,
    data_llm_alias: str = DEFAULT_DATA_LLM_ALIAS,
    code_llm_alias: str = DEFAULT_CODE_LLM_ALIAS,
    data_llm_params: Optional[Dict[str, Any]] = None,
    code_llm_params: Optional[Dict[str, Any]] = None,
):
    """Process to handle chart revision when initiated by user."""
    global chart_or_table_in_production
    agents_ready = False

    try:
        current_analysis = data_store[analysis_index]
        agentops_session_started = False
        if RUN_AGENTOPS:
            tag = ["AG2", str(os.path.basename(__file__))]
            agentops.init(api_key=os.getenv("AGENTOPS_API_KEY"), default_tags=tag)
            agentops_session_started = True

        agents_ready = initialise_agents(
            "revisechart",
            current_analysis["data_folder"],
            data_llm_alias=data_llm_alias,
            code_llm_alias=code_llm_alias,
            data_llm_params=data_llm_params,
            code_llm_params=code_llm_params,
        )

        if not agents_ready:
            current_analysis["processing"] = False
            socketio.emit("analysis_complete", {"index": analysis_index})
            if RUN_AGENTOPS and agentops_session_started:
                agentops.end_session()
            chart_or_table_in_production = ""
            return

        current_chart = current_analysis["charts_and_comments"][item_index]
        current_analysis["processing"] = True

        current_component_context = _format_chart_or_table_context(
            item=current_chart,
            analysis=current_analysis,
            include_images=True,
            include_code_token_limit=CHART_AND_TABLE_CODE_TOKEN_LIMIT,
            include_chart_csv_token_limit=CHART_DATA_REPORT_TOKEN_LIMIT,
            include_csv_filename=True,
            include_chart_image_filename=False,
        )

        other_charts_and_tables = chart_and_table_context_builder(
            analysis_index=analysis_index,
            include_images=True,
            include_code_token_limit=CHART_AND_TABLE_CODE_TOKEN_LIMIT,
            include_chart_csv_token_limit=CHART_DATA_REPORT_TOKEN_LIMIT,
            include_csv_filename=True,
            include_chart_image_filename=False,
            exclude_component_index=item_index,
        )

        # Customise process_agent system and initial message for this revision.
        process_agent_system_message = prompt_manager.get_prompt(
            "default.revise_chart.process_agent_system_message",
            user_instructions=current_analysis["instructions"],
        )

        process_agent.update_system_message(process_agent_system_message)

        additional_description = load_file_as_string(
            current_analysis["data_folder"] + "/" + "datadescription.md"
        )

        if current_chart["chart_filename"] is not None:
            process_agent_initial_message = prompt_manager.get_prompt(
                "default.revise_chart.process_agent_initial_message_chart",
                revision_instructions=instructions,
                analysis_index=analysis_index,
                item_index=item_index,
                current_chart_context=current_component_context,
                current_analysis=current_analysis,
                additional_data_description=additional_description,
                other_charts_and_tables=other_charts_and_tables,
            )
            chart_or_table_in_production = f"Chart: {current_chart['title']}"
        else:
            process_agent_initial_message = prompt_manager.get_prompt(
                "default.revise_chart.process_agent_initial_message_table",
                revision_instructions=instructions,
                analysis_index=analysis_index,
                item_index=item_index,
                current_chart_context=current_component_context,
                current_analysis=current_analysis,
                additional_data_description=additional_description,
                other_charts_and_tables=other_charts_and_tables,
            )
            chart_or_table_in_production = f"Table: {current_chart['title']}"

        result = function_agent.initiate_chat(
            main_manager,
            message=process_agent_initial_message,
            max_turns=12,
            summary_method=None,
            silent=SILENT,
        )

        filename = get_next_filename(
            os.path.join(active_analysis["data_folder"], "chat_history"),
            prefix="revise_chart_session_history_",
            suffix=".json",
        )
        with open(
            filename,
            "w",
        ) as f:
            json.dump(sanitize_for_json(result.chat_history), f, indent=4)

        current_analysis["processing"] = False
        updated_chart = current_analysis["charts_and_comments"][item_index]
        socketio.emit(
            "chart_revised",
            {
                "analysisIndex": analysis_index,
                "chartIndex": item_index,
                "updatedChart": updated_chart,
                "update_needed": True,
            },
        )
        if RUN_AGENTOPS and agentops_session_started:
            agentops.end_session()
        socketio.emit("analysis_complete", {"index": analysis_index})
        executor.stop()
        return
    except Exception as e:
        cprint(f"Error revising chart: {e}", "red")
        print(traceback.format_exc())
        current_analysis["processing"] = False
        if RUN_AGENTOPS and agentops_session_started:
            agentops.end_session()
        if agents_ready:
            executor.stop()
        socketio.emit(
            "show_popup",
            {
                "line1": "Error during analysis process, terminating.",
                "buttons": ["Ok"],
            },
        )
        wait_for_user_response()
        socketio.emit("analysis_complete", {"index": analysis_index})
        chart_or_table_in_production = ""
        return


@socketio.on("revise_chart")
def handle_revise_chart(data):
    """Handle revise chart request from frontend."""
    global active_analysis
    analysis_index = data["analysisIndex"]
    item_index = data["chartIndex"]
    instructions = data["instructions"]
    data_llm_alias = data.get("data_llm_alias", DEFAULT_DATA_LLM_ALIAS)
    code_llm_alias = data.get("code_llm_alias", DEFAULT_CODE_LLM_ALIAS)
    data_llm_params = data.get("data_llm_params", {})
    code_llm_params = data.get("code_llm_params", {})
    active_analysis = data_store[analysis_index]
    revise_chart_process(
        analysis_index,
        item_index,
        instructions,
        data_llm_alias=data_llm_alias,
        code_llm_alias=code_llm_alias,
        data_llm_params=data_llm_params,
        code_llm_params=code_llm_params,
    )
    return


def add_new_chart_process(
    analysis_index: int,
    instructions: str,
    data_llm_alias: str = DEFAULT_DATA_LLM_ALIAS,
    code_llm_alias: str = DEFAULT_CODE_LLM_ALIAS,
    data_llm_params: Optional[Dict[str, Any]] = None,
    code_llm_params: Optional[Dict[str, Any]] = None,
):
    """Process to handle adding a new chart when initiated by user."""
    global chart_or_table_in_production
    agents_ready = False

    try:
        current_analysis = data_store[analysis_index]
        agentops_session_started = False
        if RUN_AGENTOPS:
            tag = ["AG2", str(os.path.basename(__file__))]
            agentops.init(api_key=os.getenv("AGENTOPS_API_KEY"), default_tags=tag)
            agentops_session_started = True

        agents_ready = initialise_agents(
            "addnewchart",
            current_analysis["data_folder"],
            data_llm_alias=data_llm_alias,
            code_llm_alias=code_llm_alias,
            data_llm_params=data_llm_params,
            code_llm_params=code_llm_params,
        )

        if not agents_ready:
            current_analysis["processing"] = False
            socketio.emit("analysis_complete", {"index": analysis_index})
            if RUN_AGENTOPS and agentops_session_started:
                agentops.end_session()
            chart_or_table_in_production = ""
            return

        current_analysis["processing"] = True

        # Customise process_agent system message and initial for this new chart.
        process_agent_system_message = prompt_manager.get_prompt(
            "default.add_new_chart.process_agent_system_message",
            user_instructions=current_analysis["instructions"],
        )

        process_agent.update_system_message(process_agent_system_message)

        additional_description = load_file_as_string(
            current_analysis["data_folder"] + "/" + "datadescription.md"
        )

        charts_and_tables_context = chart_and_table_context_builder(
            analysis_index=analysis_index,
            include_images=True,
            include_code_token_limit=CHART_AND_TABLE_CODE_TOKEN_LIMIT,
            include_chart_csv_token_limit=CHART_DATA_REPORT_TOKEN_LIMIT,
            include_csv_filename=True,
            include_chart_image_filename=False,
        )

        process_agent_initial_message = prompt_manager.get_prompt(
            "default.add_new_chart.process_agent_initial_message",
            instructions=instructions,
            analysis_index=analysis_index,
            current_analysis=current_analysis,
            additional_data_description=additional_description,
            charts_and_tables_context=charts_and_tables_context,
        )

        result = function_agent.initiate_chat(
            main_manager,
            message=process_agent_initial_message,
            max_turns=8,
            summary_method=None,
            silent=SILENT,
        )

        filename = get_next_filename(
            os.path.join(active_analysis["data_folder"], "chat_history"),
            prefix="add_chart_session_history_",
            suffix=".json",
        )
        with open(
            filename,
            "w",
        ) as f:
            json.dump(sanitize_for_json(result.chat_history), f, indent=4)

        current_analysis["processing"] = False
        if RUN_AGENTOPS and agentops_session_started:
            agentops.end_session()
        socketio.emit("analysis_complete", {"index": analysis_index})
        executor.stop()
        return
    except Exception as e:
        cprint(f"Error adding new chart: {e}", "red")
        print(traceback.format_exc())
        current_analysis["processing"] = False
        if RUN_AGENTOPS and agentops_session_started:
            agentops.end_session()
        if agents_ready:
            executor.stop()
        socketio.emit(
            "show_popup",
            {
                "line1": "Error during analysis process, terminating.",
                "buttons": ["Ok"],
            },
        )
        wait_for_user_response()
        socketio.emit("analysis_complete", {"index": analysis_index})
        chart_or_table_in_production = ""
        return


@socketio.on("add_new_chart")
def handle_add_new_chart(data):
    """Handle add new chart request from frontend."""
    global active_analysis
    analysis_index = data["analysisIndex"]
    instructions = data["instructions"]
    data_llm_alias = data.get("data_llm_alias", DEFAULT_DATA_LLM_ALIAS)
    code_llm_alias = data.get("code_llm_alias", DEFAULT_CODE_LLM_ALIAS)
    data_llm_params = data.get("data_llm_params", {})
    code_llm_params = data.get("code_llm_params", {})
    active_analysis = data_store[analysis_index]
    add_new_chart_process(
        analysis_index,
        instructions,
        data_llm_alias=data_llm_alias,
        code_llm_alias=code_llm_alias,
        data_llm_params=data_llm_params,
        code_llm_params=code_llm_params,
    )
    return


@socketio.on("revise_summary")
def handle_revise_summary(data):
    """Handle revise summary request from frontend."""
    global active_analysis
    try:
        analysis_index = data["analysisIndex"]
        active_analysis = data_store[analysis_index]
        analysis = data_store[analysis_index]
        instructions = data["instructions"]
        include_images_in_html = data.get("include_images", False)
        if not ALLOW_HTML_SUMMARY_GENERATION:
            include_images_in_html = False
        summary_llm_alias = data.get("summary_llm_alias", DEFAULT_SUMMARY_LLM_ALIAS)
        summary_llm_params = data.get("summary_llm_params", {})
        analysis["processing"] = True
        summary_data = create_summary(
            analysis,
            instructions,
            include_images_in_html=include_images_in_html,
            summary_llm_alias=summary_llm_alias,
            summary_llm_params=summary_llm_params,
        )
        socketio.emit(
            "summary_update",
            {"analysisIndex": analysis_index, "summary": summary_data},
        )
    except Exception as e:
        cprint(f"Error generating summary: {e}", "red", attrs=["bold"])
        traceback.print_exc()
        socketio.emit(
            "show_popup",
            {
                "line1": "Error generating summary.",
                "buttons": ["Ok"],
            },
        )
        wait_for_user_response()
    analysis["processing"] = False
    socketio.emit("analysis_complete", {"index": analysis_index})
    return


@socketio.on("open_folder")
def handle_open_folder(data):
    """
    Open the requested folder in the OS file explorer.
    Expects: {"folder": "data_store/abcd12"}
    """
    rel_path = data.get("folder", "")
    abs_path = os.path.abspath(rel_path)

    if not os.path.isdir(abs_path):
        cprint(f"[open_folder] Path not found: {abs_path}", "red")
        return

    try:
        if platform.system() == "Windows":
            os.startfile(abs_path)  # type: ignore
        elif platform.system() == "Darwin":  # macOS
            subprocess.Popen(["open", abs_path])
        else:  # Linux / WSL
            subprocess.Popen(["xdg-open", abs_path])
        print(f"[open_folder] Opened {abs_path}")
    except Exception as exc:
        cprint(f"[open_folder] Failed to open {abs_path}: {exc}", "red")


@socketio.on("revise_report")
def handle_revise_report(data):
    """Handle revise report request from frontend."""
    global active_analysis, llm_report_review
    try:
        executor = None
        analysis_index = data["analysisIndex"]
        active_analysis = data_store[analysis_index]
        analysis = data_store[analysis_index]
        instructions = data["instructions"]
        include_existing_report = data["include_context"]
        llm_alias = data.get("llm_alias", DEFAULT_REPORT_LLM_ALIAS)
        llm_report_review = data.get("llm_report_review", False)
        llm_params = data.get("llm_params", {})
        analysis["processing"] = True
        executor, docker_error = docker_start(active_analysis["data_folder"])
        if docker_error:
            analysis["processing"] = False
            cprint("Docker not available, terminating.\n\n", "red", attrs=["bold"])
            socketio.emit(
                "show_popup",
                {
                    "line1": "Docker not available, terminating.",
                    "buttons": ["Ok"],
                },
            )
            wait_for_user_response()
            socketio.emit("analysis_complete", {"index": analysis_index})
            return

        generate_report_success = generate_report(
            analysis,
            instructions,
            executor,
            include_existing_report,
            llm_alias,
            llm_params=llm_params,
        )
        if generate_report_success is False:
            cprint("Error generating report", "red", attrs=["bold"])
            socketio.emit(
                "show_popup",
                {
                    "line1": "Error generating report.",
                    "buttons": ["Ok"],
                },
            )
            wait_for_user_response()

    except Exception as e:
        cprint(f"Error generating report: {e}", "red", attrs=["bold"])
        if executor is not None:
            executor.stop()
        traceback.print_exc()
        socketio.emit(
            "show_popup",
            {
                "line1": "Error generating report.",
                "buttons": ["Ok"],
            },
        )
        wait_for_user_response()
        analysis["processing"] = False
        socketio.emit("analysis_complete", {"index": analysis_index})
        return

    socketio.emit("report_update", generate_report_success)
    analysis["processing"] = False
    socketio.emit("analysis_complete", {"index": analysis_index})
    return


def build_report_message(analysis, instructions, include_existing_report):
    if instructions == "":
        instructions = "Please write a report based on the analysis provided."

    try:
        analysis_index = data_store.index(analysis)
    except ValueError as e:
        for idx, stored_analysis in enumerate(data_store):
            if stored_analysis.get("data_folder") == analysis.get("data_folder"):
                analysis_index = idx
                break
        else:
            raise e

    if analysis["report"]["report_exists"] and include_existing_report:
        existing_report = (
            "Here is the existing report for revision: \n"
            + generate_png_list_string(analysis["data_folder"] + "/report_images")
        )
        existing_report += (
            "\nThe existing report was created from this code: \n"
            + load_file_as_string(os.path.join(analysis["data_folder"], "report.py"))
        )
    else:
        existing_report = ""

    additional_data_description = load_file_as_string(
        analysis["data_folder"] + "/" + "datadescription.md"
    )

    charts_and_tables = chart_and_table_context_builder(
        analysis_index=analysis_index,
        include_images=True,
        include_code_token_limit=None,
        include_chart_csv_token_limit=CHART_DATA_REPORT_TOKEN_LIMIT,
        include_csv_filename=False,
        include_chart_image_filename=True,
    )

    report_message = prompt_manager.get_prompt(
        "default.report.report_agent_initial_message",
        instructions=instructions,
        user_instructions=analysis["instructions"],
        existing_report=existing_report,
        current_analysis=analysis,
        additional_data_description=additional_data_description,
        charts_and_tables=charts_and_tables,
    )

    return report_message


def generate_report(
    analysis,
    instructions,
    executor,
    include_existing_report,
    llm_alias=DEFAULT_REPORT_LLM_ALIAS,
    llm_params: Optional[Dict[str, Any]] = None,
):
    """Generate a report from the analysis data and instructions."""
    try:
        report_reply = None
        message = "Generating report..."
        cprint(message, "cyan", attrs=["bold"])
        socketio.emit("update_popup", {"message": message})

        agentops_session_started = False
        if RUN_AGENTOPS:
            tag = ["AG2", str(os.path.basename(__file__))]
            agentops.init(api_key=os.getenv("AGENTOPS_API_KEY"), default_tags=tag)
            agentops_session_started = True

        if "report" not in analysis:
            analysis["report"] = {"report_exists": False, "update_needed": True}

        report_template_code = load_file_as_string("prompts/report_template.py")
        report_agent_sys_message = prompt_manager.get_prompt(
            "default.report.report_agent_sys_message",
            report_template_code=report_template_code,
        )

        report_message = build_report_message(
            analysis, instructions, include_existing_report
        )

        cprint(f"Report LLM:\n{cfg(llm_alias, **(llm_params or {}))}\n", "cyan")

        cprint(
            f"report_agent system message tokens: {autogen.token_count_utils.count_token(report_agent_sys_message, model='gpt-4o')}",
            "cyan",
        )
        cprint(
            f"report message tokens: {autogen.token_count_utils.count_token(report_message, model='gpt-4o')}",
            "cyan",
        )

        report_agent = MultimodalConversableAgent(
            name="report_agent",
            system_message=report_agent_sys_message,
            max_consecutive_auto_reply=5,
            code_execution_config=False,
            human_input_mode="NEVER",
            is_termination_msg=lambda msg: (
                "terminate"
                in (
                    msg["content"][0]["text"].lower()
                    if isinstance(msg.get("content"), list)
                    and len(msg["content"]) > 0
                    and isinstance(msg["content"][0], dict)
                    and "text" in msg["content"][0]
                    else msg["content"].lower()
                    if isinstance(msg.get("content"), str)
                    else ""
                )
            ),
            llm_config={
                "config_list": cfg(llm_alias, **(llm_params or {})),
                "cache_seed": AGENT_CACHE,
            },
        )

        report_code_interpreter = autogen.UserProxyAgent(
            "report_code_interpreter",
            human_input_mode="NEVER",
            code_execution_config={"executor": executor},
            is_termination_msg=lambda msg: "task complete" in msg["content"].lower(),
            llm_config=False,
            default_auto_reply="",
        )

        report_code_interpreter.register_hook(
            hookable_method="process_message_before_send",
            hook=hook_return_report_to_report_agent,
        )

        # Make sure any existing libreoffice report Docker containers are removed before starting
        if process_docx_method == "libreoffice":
            stop_lo_container(remove=True)

        cprint("Calling report writing agent...", "cyan")
        report_reply = report_code_interpreter.initiate_chat(
            report_agent,
            message=report_message,
            silent=SILENT,
        )

        filename = get_next_filename(
            os.path.join(active_analysis["data_folder"], "chat_history"),
            prefix="report_session_history_",
            suffix=".json",
        )
        with open(
            filename,
            "w",
        ) as f:
            json.dump(sanitize_for_json(report_reply.chat_history), f, indent=4)

        executor.stop()
        executor = None
        if process_docx_method == "libreoffice":
            stop_lo_container(remove=True)

        if RUN_AGENTOPS and agentops_session_started:
            agentops.end_session()
        # TODO: implement improved logic to check if report successfully produced
        if (
            "TASK COMPLETE" in report_reply.chat_history[-1]["content"].upper()
            or "TERMINATE" in report_reply.chat_history[-1]["content"].upper()
        ) and os.path.isfile(f"{analysis['data_folder']}/report.docx"):
            cprint("Report received.", "cyan")
            analysis["report"]["report_exists"] = True
            analysis["report"]["update_needed"] = False
            save_data_store()
            return True
        else:
            cprint("Report generation failed.", "red")
            analysis["report"]["report_exists"] = False
            analysis["report"]["update_needed"] = True
            save_data_store()
            return False
    except Exception as e:
        cprint(f"Error in generate_report: {e}", "red", attrs=["bold"])
        if report_reply is not None:
            filename = get_next_filename(
                os.path.join(active_analysis["data_folder"], "chat_history"),
                prefix="report_session_history_",
                suffix=".json",
            )
            with open(
                filename,
                "w",
            ) as f:
                json.dump(sanitize_for_json(report_reply.chat_history), f, indent=4)
        if executor is not None:
            executor.stop()
        if process_docx_method == "libreoffice":
            stop_lo_container(remove=True)
        traceback.print_exc()
        analysis["report"]["report_exists"] = False
        analysis["report"]["update_needed"] = True
        save_data_store()
        return False


def hook_return_report_to_report_agent(
    sender: ConversableAgent,
    message: Union[dict[str, Any], str],
    recipient: Agent,
    silent: bool,
) -> Union[dict[str, Any], str]:
    """If successful execution, process the word document and get review message."""

    if isinstance(message, str):
        if "exitcode: 0 (execution succeeded)" in message:
            # check report.docx exists
            report_path = os.path.join(active_analysis["data_folder"], "report.docx")
            if os.path.exists(report_path):
                processing_result = process_docx(
                    report_path,
                    active_analysis["data_folder"],
                    dpi=150,
                )
                if processing_result:
                    message = report_review_message()
                else:  # report exists but failed processing - terminate
                    cprint(
                        "Docx report file created, but error processing. Terminating report production session.",
                        "red",
                    )
                    message = "Terminate"
            else:
                cprint(
                    "Provided script executed successfully but did not create report.docx file.",
                    "cyan",
                )
                message = (
                    "Provided script executed successfully but did not create report.docx file. \nScript output: \n"
                    + message
                )

    return message


def report_review_message() -> str:
    """Create a message to review the report."""

    if not llm_report_review:
        return "Terminate"

    status_message = "Reviewing report..."
    cprint(status_message, "cyan", attrs=["bold"])
    socketio.emit("update_popup", {"message": status_message})

    if process_docx_method == "word":
        message = prompt_manager.get_prompt(
            "default.report.report_agent_review_message",
            report=generate_png_list_string(
                active_analysis["data_folder"] + "/report_images"
            ),
        )
    elif process_docx_method == "libreoffice":
        message = prompt_manager.get_prompt(
            "default.report.report_agent_review_message_no_fields",
            report=generate_png_list_string(
                active_analysis["data_folder"] + "/report_images"
            ),
        )

    return message


def process_docx(docx_path, output_folder, dpi=150) -> bool:
    """Process the DOCX file to refresh fields, generate a PDF and convert to PNG images."""

    cprint(f"Processing document: {docx_path}", "cyan")

    status_message = "Processing report..."
    cprint(status_message, "cyan", attrs=["bold"])
    socketio.emit("update_popup", {"message": status_message})

    # Input validation
    if not docx_path or not os.path.exists(docx_path):
        cprint(f"Error: Document file not found: {docx_path}", "red")
        return False

    if not output_folder:
        cprint("Error: Output folder path cannot be empty", "red")
        return False

    if process_docx_method is None:
        cprint(
            "DOCX processing method is not configured because the required tooling is unavailable.",
            "red",
        )
        return False

    if process_docx_method == "word":
        processed = process_docx_via_word(docx_path, output_folder)
    elif process_docx_method == "libreoffice":
        processed = process_docx_via_libreoffice(
            docx_path=docx_path,
            output_folder=output_folder,
        )
    else:
        cprint(f"Unsupported DOCX processing method '{process_docx_method}'.", "red")
        return False

    if not processed:
        cprint("Error processing DOCX file.", "red")
        return False

    scrub_docx_metadata(
        in_docx=docx_path,
        overwrite=True,
        remove_timestamps=False,
        remove_custom_properties_entirely=True,
    )
    cprint("Scrubbed DOCX metadata.", "cyan")
    base_filename = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_folder, f"{base_filename}.pdf")
    scrub_pdf_metadata(in_pdf=pdf_path, overwrite=True)
    cprint("Scrubbed PDF metadata.", "cyan")

    if convert_pdf_to_images(docx_path, output_folder, dpi=dpi):
        cprint("Document processing complete.", "cyan")
        return True
    else:
        cprint("Error converting PDF to images.", "red")
        return False


def convert_pdf_to_images(docx_path, output_folder, dpi=150) -> bool:
    """Convert the PDF created from the DOCX to PNG images."""
    pdf = None
    try:
        images_folder = os.path.join(output_folder, "report_images")
        os.makedirs(images_folder, exist_ok=True)
    except OSError as e:
        cprint(f"Error creating image folder: {e}", "red")
        return False

    cprint("Removing existing page images...", "cyan")
    page_pattern = compile(r"^page_\d+\.png$")
    try:
        for filename in os.listdir(images_folder):
            if page_pattern.match(filename):
                file_path = os.path.join(images_folder, filename)
                try:
                    os.remove(file_path)
                    cprint(f"Deleted: {filename}", "cyan")
                except Exception as e:
                    cprint(f"Warning: Error deleting {filename}: {e}", "red")
    except Exception as e:
        cprint(f"Warning: Error listing files in images folder: {e}", "red")

    base_filename = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_folder, f"{base_filename}.pdf")
    cprint("Extracting images from PDF...", "cyan")
    try:
        pdf = fitz.open(pdf_path)
        zoom = dpi / 72
        matrix = fitz.Matrix(zoom, zoom)
        page_count = len(pdf)
        if page_count == 0:
            cprint("Warning: PDF has no pages", "red")
            return False

        for page_num, page in enumerate(pdf):
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            output_file = os.path.join(images_folder, f"page_{page_num + 1}.png")
            pix.save(output_file)
            cprint(f"Saved page {page_num + 1} as {output_file}", "cyan")

        pdf.close()
        pdf = None  # Mark as closed
        return True
    except Exception as e:
        cprint(f"Error processing PDF: {e}", "red")
        return False
    finally:
        try:
            if pdf:
                pdf.close()
        except Exception as e:
            cprint(f"Warning: Error closing PDF: {e}", "red")


def process_docx_via_word(docx_path, output_folder):
    """Process DOCX using Microsoft Word via COM."""

    pythoncom.CoInitialize()
    word_app = None
    doc = None

    try:
        try:
            os.makedirs(output_folder, exist_ok=True)
        except OSError as e:
            cprint(f"Error creating output folder: {e}", "red")
            return False

        base_filename = os.path.splitext(os.path.basename(docx_path))[0]
        pdf_path = os.path.join(output_folder, f"{base_filename}.pdf")

        # Create a new, separate Word instance
        try:
            word_app = win32.DispatchEx("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False
            cprint("Created new separate Word instance", "cyan")
        except Exception as e:
            cprint(f"Error creating Word application: {e}", "red")
            return False

        abs_path = os.path.abspath(docx_path)

        cprint("Updating Word fields...", "cyan")
        try:
            doc = word_app.Documents.Open(
                abs_path, ReadOnly=False, Visible=False, ConfirmConversions=False
            )
        except Exception as e:
            cprint(f"Error opening Word document: {e}", "red")
            return False

        try:
            # Update all fields in the document
            for i in range(2):  # Sometimes needs multiple passes
                try:
                    doc.Fields.Update()
                except Exception as e:
                    cprint(f"Warning: Error updating fields (pass {i + 1}): {e}", "red")

                # Update all tables of contents
                try:
                    if doc.TablesOfContents.Count > 0:
                        for toc in doc.TablesOfContents:
                            toc.Update()
                except Exception as e:
                    cprint(f"Warning: Error updating tables of contents: {e}", "red")

                # Update all tables of figures
                try:
                    if (
                        hasattr(doc, "TablesOfFigures")
                        and doc.TablesOfFigures.Count > 0
                    ):
                        for tof in doc.TablesOfFigures:
                            tof.Update()
                except Exception as e:
                    cprint(f"Warning: Error updating tables of figures: {e}", "red")

            # Save changes after updating fields
            doc.Save()
            cprint("Fields updated successfully", "cyan")
        except Exception as e:
            cprint(f"Error during field updates: {e}", "red")
            # Continue processing even if field updates fail

        # Convert to PDF
        cprint(f"Converting to PDF: {pdf_path}", "cyan")
        try:
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)  # 17 = PDF format
            doc.Close(SaveChanges=True)
            doc = None  # Mark as closed
            cprint(f"Converted DOCX to PDF: {pdf_path}", "cyan")
        except Exception as e:
            cprint(f"Error converting to PDF: {e}", "red")
            return False

        # Verify PDF was created
        if not os.path.exists(pdf_path):
            cprint(f"Error: PDF file was not created: {pdf_path}", "red")
            return False

        cprint(f"Successfully processed {docx_path}", "cyan")
        return True

    except Exception as e:
        cprint(f"Unexpected error processing {docx_path}: {e}", "red")
        print(traceback.format_exc())
        return False

    finally:
        # Cleanup resources in reverse order
        try:
            if doc:
                doc.Close(SaveChanges=False)
        except Exception as e:
            cprint(f"Warning: Error closing Word document: {e}", "red")

        try:
            if word_app:
                word_app.Quit()
                time.sleep(1)  # Give Word time to fully close
        except Exception as e:
            cprint(f"Warning: Error closing Word application: {e}", "red")

        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            cprint(f"Warning: Error uninitializing COM: {e}", "red")


def generate_png_list_string(folder_path):

    pattern = compile(r"^page_(\d+)\.png$")

    matched_files = []

    # List all files in the directory
    for filename in os.listdir(folder_path):
        match = pattern.match(filename)
        if match:
            # Store as tuple (number, filename) for sorting
            matched_files.append((int(match.group(1)), filename))

    # Sort by numeric value
    matched_files.sort()

    result = "\n".join(
        f"<img {folder_path}/{filename}>" for _, filename in matched_files
    )
    return result


def create_summary(
    analysis,
    instructions,
    include_images_in_html: bool = False,
    summary_llm_alias: str = DEFAULT_SUMMARY_LLM_ALIAS,
    summary_llm_params: Optional[Dict[str, Any]] = None,
):
    """
    Run data through the summary process.
    Returns: str output from LLM summariser.
    """
    include_images_in_html = include_images_in_html and ALLOW_HTML_SUMMARY_GENERATION

    message = "Summarising analysis..."
    cprint(message, "cyan", attrs=["bold"])
    socketio.emit("update_popup", {"message": message})

    cprint(
        f"Summary LLM:\n{cfg(summary_llm_alias, **(summary_llm_params or {}))}\n",
        "cyan",
    )

    try:
        analysis_index = data_store.index(analysis)
    except ValueError as e:
        for idx, stored_analysis in enumerate(data_store):
            if stored_analysis.get("data_folder") == analysis.get("data_folder"):
                analysis_index = idx
                break
        else:
            raise e

    sheet_details = ""
    for itemsheets in analysis["description"]["sheets"]:
        sheet_details += f"""Sheet name: {itemsheets["sheet_name"]}
Headings: {", ".join(map(str, itemsheets["headings"]))}
Time period: {itemsheets["time_start"]} to {itemsheets["time_end"]}
Time step: {itemsheets["time_step"]}
Time format: {itemsheets["time_format"]}

"""

    comments_for_summary = (
        f"Data description: {analysis['description']['data_description']}\n\n"
        + sheet_details
        + chart_and_table_context_builder(
            analysis_index=analysis_index,
            include_images=True,
            include_code_token_limit=None,
            include_chart_csv_token_limit=CHART_DATA_REPORT_TOKEN_LIMIT,
            include_csv_filename=False,
            include_chart_image_filename=include_images_in_html,
        )
    )
    prompt_key = (
        "default.summary.summary_agent_sys_message_html"
        if include_images_in_html
        else "default.summary.summary_agent_sys_message_md"
    )
    summary_agent_sys_message = prompt_manager.get_prompt(prompt_key)

    summary_agent = MultimodalConversableAgent(
        name="summary_agent",
        system_message=summary_agent_sys_message,
        max_consecutive_auto_reply=1,
        code_execution_config=False,
        human_input_mode="NEVER",
        llm_config={
            "config_list": cfg(summary_llm_alias, **(summary_llm_params or {})),
            "timeout": 40,
        },
    )

    if instructions == "":
        summary_message = prompt_manager.get_prompt(
            "default.summary.summary_message_without_instructions",
            user_instructions=user_instructions,
            comments_for_summary=comments_for_summary,
        )
    else:
        summary_message = prompt_manager.get_prompt(
            "default.summary.summary_message_with_instructions",
            user_instructions=user_instructions,
            instructions=instructions,
            comments_for_summary=comments_for_summary,
        )

    user_proxy = autogen.UserProxyAgent(
        "user_proxy",
        human_input_mode="NEVER",
        max_consecutive_auto_reply=0,
        code_execution_config=False,
    )

    cprint("Calling summary agent...", "cyan")
    summary_reply = user_proxy.initiate_chat(
        summary_agent,
        message=summary_message,
        silent=SILENT,
    )
    cprint("Summary received.", "cyan")

    folder = os.path.normpath(
        os.path.join(active_analysis["data_folder"], "chat_history")
    )
    os.makedirs(folder, exist_ok=True)

    # Save chat_history to file
    filename = get_next_filename(
        folder,
        prefix="summary_session_history_",
        suffix=".json",
    )
    with open(
        filename,
        "w",
    ) as f:
        json.dump(sanitize_for_json(summary_reply.chat_history), f, indent=4)

    summary_reply = summary_reply.chat_history[1]["content"]
    if isinstance(summary_reply, str):
        lines = summary_reply.split("\n")
    elif isinstance(summary_reply, dict):
        lines = summary_reply["content"].split("\n")
    # Check if the first line is "Summary of findings" (case insensitive)
    if lines and "summary of findings" in lines[0].strip().lower():
        # Remove the first line
        lines = lines[1:]
        # Join the remaining lines back into a single string
        summary_reply = "\n".join(lines)

    if include_images_in_html:
        summary_reply = sanitise_html(summary_reply)

    summary_data = analysis["summary"]
    summary_data["summary_comments"] = summary_reply
    summary_data["update_needed"] = False
    save_data_store()
    return summary_data


def generate_random_folder_name(length=6):
    """Create a new folder with a random name to store analysis."""
    # Define the characters to use (alphanumeric only)
    characters = string.ascii_letters + string.digits
    # Generate a random string of the specified length
    folder_name = "".join(random.choices(characters, k=length))
    folder_path = os.path.join(app.config["UPLOAD_FOLDER"], folder_name)
    try:
        os.mkdir(folder_path)
        print(f"Folder created: {folder_path}")
    except FileExistsError:
        cprint(f"Folder already exists: {folder_path}", "red")
    except OSError as e:
        cprint(f"Error creating folder: {e}", "red")
    return folder_path


def safe_read_csv(path, **kwargs):
    """
    Robust CSV reader that handles UTF-8, Windows encodings, BOM, and UTF-16.
    """
    path = Path(path)

    # --- Detect UTF-16 by checking file size vs byte length ---
    try:
        size = path.stat().st_size
        # UTF-16: every character is 2 bytes, file often contains many 0x00 bytes
        looks_utf16 = size > 0 and size % 2 == 0
    except Exception:
        looks_utf16 = False

    # --- Encodings to try ---
    encodings = [
        "utf-8",
        "utf-8-sig",  # BOM-safe
        "cp1252",  # most common Windows CSV encoding
        "latin1",
        "iso-8859-1",
    ]

    # Try standard encodings first
    last_err = None
    for enc in encodings:
        try:
            return pd.read_csv(path, encoding=enc, **kwargs), enc
        except UnicodeDecodeError as e:
            last_err = e

    # Try UTF-16 **only if** file probably matches pattern
    if looks_utf16:
        for enc in ["utf-16", "utf-16-le", "utf-16-be"]:
            try:
                return pd.read_csv(path, encoding=enc, **kwargs), enc
            except Exception:
                pass  # try next

    # Final fallback: no decoding crash but replace invalid chars
    try:
        return pd.read_csv(
            path, encoding="latin1", encoding_errors="replace", **kwargs
        ), "latin1 (with replacement)"
    except Exception as e:
        raise last_err or e


def test_generate_data_description(file_name) -> tuple[str, str]:
    """Hard coded test run of the data description function in prompt default.initial_data_investigation_large, except file is not created.
    This is used to determine if the data description is too many tokens.
    Update to match prompt if prompt is changed."""
    file_name = file_name
    csv_encoding = None

    ext = Path(file_name).suffix

    # Build a dict of {sheet_name: DataFrame}
    if ext in {".xlsx", ".xls"}:
        xls = pd.ExcelFile(file_name)
        data_frames = {
            sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names
        }
    elif ext == ".csv":
        df_csv, csv_encoding = safe_read_csv(file_name)
        sheet_name = "CSV Sheet"
        data_frames = {sheet_name: df_csv}
    else:
        raise ValueError(f"Unsupported file extension ‘{ext}’.")

    datadescription = ""

    try:
        datadescription += f"# Data description for {file_name}\n\n"

        if ext == ".csv":
            datadescription += f"## CSV Encoding: The expected CSV encoding to use with this CSV is {csv_encoding}\n\n"

        if ext in {".xlsx", ".xls"}:
            sheet_count = len(data_frames)
            datadescription += f"## Number of sheets: \n{sheet_count}\n\n"

        for sheet_name, df in data_frames.items():
            datadescription += f"## Sheet name: {sheet_name}\n\n"

            # Column headings
            datadescription += "### Column Headings\n"
            datadescription += ", ".join([f"'{col}'" for col in df.columns]) + "\n\n"

            # First 10 rows
            datadescription += "### First 10 Rows\n"
            datadescription += df.head(5).to_markdown() + "\n\n"

            # Bottom 10 rows
            datadescription += "### Bottom 10 Rows\n"
            datadescription += df.tail(10).to_markdown() + "\n\n"

            # Describe function include='all'
            datadescription += "### Describe (include='all')\n"
            datadescription += df.describe(include="all").to_markdown() + "\n\n"

            # Info function verbose=True
            datadescription += "### Info (verbose=True)\n"
            buffer = io.StringIO()
            df.info(buf=buffer, verbose=True)
            info_output = buffer.getvalue()
            datadescription += "```\n" + info_output + "\n```\n\n"

            # Shape function
            datadescription += "### Shape\n"
            datadescription += str(df.shape) + "\n\n"

            # Text columns unique values (excluding numeric values within the column, dates, and columns with 'date' in name)
            text_columns = [
                col
                for col in df.columns
                if df[col].dtype == object and "date" not in col.lower()
            ]
            if text_columns:
                datadescription += "### Unique values in text columns (non-numeric only, up to 100 values)\n"
                for col in text_columns:
                    series = df[col]
                    has_nan = series.isna().any()
                    non_null = series.dropna()
                    stripped = non_null.astype(str).str.strip()
                    has_empty = (stripped == "").any()

                    non_empty = non_null[stripped != ""]

                    as_str = non_empty.astype(str).str.strip()
                    numeric_mask = pd.to_numeric(as_str, errors="coerce").notna()
                    non_numeric = non_empty[~numeric_mask]

                    uniques = non_numeric.unique().tolist()

                    if has_empty:
                        uniques.append("<empty string>")
                    if has_nan:
                        uniques.append("<NaN>")

                    unique_vals_quoted = [f"'{str(val)}'" for val in uniques[:100]]
                    datadescription += f"#### Column: {col}\n"
                    datadescription += ", ".join(unique_vals_quoted) + "\n\n"
    except Exception as e:
        cprint(f"Error generating data description: {e}", "red")
        return "use small", csv_encoding
    data_description_token_count = autogen.token_count_utils.count_token(
        datadescription, model="gpt-4o"
    )

    if data_description_token_count > DATA_DESCRIPTION_LARGE_TOKEN_LIMIT:
        print(f"Data description too long: {data_description_token_count} tokens.")
        return "use small", csv_encoding
    else:
        print(
            f"Data description generated successfully: {data_description_token_count} tokens."
        )
        return "use large", csv_encoding


def sanitize_for_json(obj):
    """Recursively replace non-serializable objects (like Images) with placeholders."""
    if isinstance(obj, dict):
        return {k: sanitize_for_json(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [sanitize_for_json(v) for v in obj]
    elif isinstance(obj, Image.Image):  # check if it's a PIL Image
        return "image content removed"
    else:
        return obj


@app.route("/")
def index():
    html_file_name = HTML_INTERFACE_FILENAME
    return render_template(
        html_file_name,
        data_store=data_store,
        llm_aliases=ALIASES,
        data_llm_aliases=ALIASES_VISION,
        code_llm_aliases=ALIASES,
        summary_llm_aliases=ALIASES_VISION,
        default_report_llm_alias=DEFAULT_REPORT_LLM_ALIAS,
        default_data_llm_alias=DEFAULT_DATA_LLM_ALIAS,
        default_code_llm_alias=DEFAULT_CODE_LLM_ALIAS,
        default_summary_llm_alias=DEFAULT_SUMMARY_LLM_ALIAS,
        model_info=MODEL_INFO,
        anthropic_thinking_enable=ANTHROPIC_THINKING_ENABLE,
        allow_html_summary_generation=ALLOW_HTML_SUMMARY_GENERATION,
        llm_api_key_status=API_KEY_STATUS,
        process_docx_method=process_docx_method,
        table_sig_figs=TABLE_SIG_FIGS,
        rounding_exclude_keywords=EXCLUDE_KEYWORDS,
        rounding_override_include_keywords=OVERRIDE_INCLUDE_KEYWORDS,
    )


@app.route("/report_viewer")
def report_viewer():
    return render_template("report_pdf_viewer.html")


@app.route("/startanalysis", methods=["POST"])
def run_analysis():
    """Run the analysis process."""
    global user_instructions, active_analysis, chart_or_table_in_production
    active_folder = generate_random_folder_name()
    active_folder = active_folder.replace("\\", "/")
    analysis_history_to_be_saved = False
    analysis_history_saved = False
    agentops_session_started = False

    if "file" not in request.files:
        return redirect(request.url)
    file = request.files["file"]
    if file.filename == "":
        return redirect(request.url)
    if file and allowed_file(file.filename):
        try:
            filename = secure_filename(file.filename)
            file_path = os.path.normpath(os.path.join(active_folder, filename))
            file.save(file_path)
            initial_message = request.form["instructions"]
            user_instructions = initial_message

            data_llm_alias = request.form.get("data_llm_alias", DEFAULT_DATA_LLM_ALIAS)
            code_llm_alias = request.form.get("code_llm_alias", DEFAULT_CODE_LLM_ALIAS)
            summary_llm_alias = request.form.get(
                "summary_llm_alias", DEFAULT_SUMMARY_LLM_ALIAS
            )
            data_llm_params = json.loads(
                request.form.get("data_llm_params", "{}") or "{}"
            )
            code_llm_params = json.loads(
                request.form.get("code_llm_params", "{}") or "{}"
            )
            summary_llm_params = json.loads(
                request.form.get("summary_llm_params", "{}") or "{}"
            )

            new_analysis = {
                "data_file_name": filename,
                "data_folder": active_folder,
                "instructions": initial_message,
                "user_notes": DEFAULT_USER_NOTES,
                "date": datetime.now().strftime("%Y-%m-%d"),
                "name": f"Analysis {len(data_store) + 1}",
                "description": {},
                "charts_and_comments": [],
                "summary": {"summary_comments": "", "update_needed": True},
                "report": {"report_exists": False, "update_needed": True},
                "processing": True,
            }

            active_analysis = new_analysis

            folder_path = os.path.join(active_analysis["data_folder"], "chat_history")
            os.makedirs(folder_path, exist_ok=True)

            if RUN_AGENTOPS:
                tag = ["AG2", str(os.path.basename(__file__))]
                agentops.init(api_key=os.getenv("AGENTOPS_API_KEY"), default_tags=tag)
                agentops_session_started = True

            agents_ready = initialise_agents(
                "mainprocess",
                active_folder,
                data_llm_alias=data_llm_alias,
                code_llm_alias=code_llm_alias,
                data_llm_params=data_llm_params,
                code_llm_params=code_llm_params,
            )

            if not agents_ready:
                new_analysis["processing"] = False
                if new_analysis in data_store:
                    data_store.remove(new_analysis)
                    save_data_store()
                if RUN_AGENTOPS and agentops_session_started:
                    agentops.end_session()
                chart_or_table_in_production = ""
                shutil.rmtree(active_folder, ignore_errors=True)
                active_analysis = []
                return "", 500

        except Exception as e:
            cprint(f"Error during analysis setup: {e}", "red", attrs=["bold"])
            traceback.print_exc()
            if new_analysis in data_store:
                data_store.remove(new_analysis)
                save_data_store()
            if RUN_AGENTOPS and agentops_session_started:
                agentops.end_session()
            chart_or_table_in_production = ""
            shutil.rmtree(active_folder, ignore_errors=True)
            active_analysis = []
            socketio.emit(
                "show_popup",
                {
                    "line1": "Error during analysis setup, terminating.",
                    "buttons": ["Ok"],
                },
            )
            wait_for_user_response()
            return "", 500

        try:
            data_store.append(new_analysis)
            save_data_store()

            socketio.emit(
                "analysis_start",
                {
                    "data_file_name": filename,
                    "data_folder": active_folder,
                    "instructions": initial_message,
                    "user_notes": DEFAULT_USER_NOTES,
                    "date": new_analysis["date"],
                    "name": new_analysis["name"],
                    "summary": new_analysis["summary"],
                    "processing": True,
                },
            )

            ###Step to replace mu and mirco with 'u' to avoid error that sometimes occurs with 'Âµg/L' and 'Î¼'. Error does not always occur. Disable step for now.
            # cprint(f"Sanitizing micro units in file: {file_path}", "cyan")
            # from utils.sanitize_mu_units import sanitize_mu_micro_inplace

            # result = sanitize_mu_micro_inplace(file_path)
            # cprint(f"Sanitization result: {result}", "cyan")

            data_description_size, csv_encoding = test_generate_data_description(
                file_path
            )

            if csv_encoding is not None:
                csv_encoding_text = (
                    f"The expected CSV encoding to use with this CSV is {csv_encoding}."
                )
                if csv_encoding == "latin1 (with replacement)":
                    csv_encoding = 'encoding="latin1", encoding_errors="replace"'
                else:
                    csv_encoding = f'encoding="{csv_encoding}"'
            else:
                csv_encoding_text = ""
                csv_encoding = ""

            if data_description_size == "use large":
                start_prompt = prompt_manager.get_prompt(
                    "default.initial_data_investigation_large",
                    filename=filename,
                    csv_encoding=csv_encoding,
                    csv_encoding_text=csv_encoding_text,
                )
            elif data_description_size == "use small":
                start_prompt = prompt_manager.get_prompt(
                    "default.initial_data_investigation_small",
                    filename=filename,
                    csv_encoding=csv_encoding,
                    csv_encoding_text=csv_encoding_text,
                )

            chart_or_table_in_production = "initial data review"

            analysis_history_to_be_saved = True

            main_process_result = process_agent.initiate_chat(
                main_manager,
                message=start_prompt,
                max_turns=16,
                summary_method=None,
                silent=SILENT,
            )

            with open(
                os.path.join(active_folder, "chat_history/main_process_history.json"),
                "w",
            ) as f:
                json.dump(
                    sanitize_for_json(main_process_result.chat_history), f, indent=4
                )
                analysis_history_saved = True

            summary_data = create_summary(
                analysis=new_analysis,
                instructions="",
                include_images_in_html=False,
                summary_llm_alias=summary_llm_alias,
                summary_llm_params=summary_llm_params,
            )
            socketio.emit(
                "summary_update",
                {"analysisIndex": len(data_store) - 1, "summary": summary_data},
            )

            message = "Cleaning up..."
            cprint(message, "cyan", attrs=["bold"])
            socketio.emit("update_popup", {"line1": message})

            new_analysis["processing"] = False
            chart_or_table_in_production = ""
            executor.stop()
            if RUN_AGENTOPS and agentops_session_started:
                agentops.end_session()

            socketio.emit("analysis_complete", {"index": len(data_store) - 1})
            return "", 204  # No Content, prevents page refresh

        except Exception as e:
            cprint(f"Error during analysis process: {e}", "red", attrs=["bold"])
            traceback.print_exc()
            new_analysis["processing"] = False
            chart_or_table_in_production = ""
            save_data_store()
            if agents_ready:
                executor.stop()

            if analysis_history_to_be_saved and not analysis_history_saved:
                try:
                    with open(
                        os.path.join(
                            active_folder, "chat_history/main_process_history.json"
                        ),
                        "w",
                    ) as f:
                        json.dump(
                            sanitize_for_json(main_process_result.chat_history),
                            f,
                            indent=4,
                        )
                except Exception as e:
                    cprint(
                        f"Error saving analysis history after failure: {e}",
                        "red",
                        attrs=["bold"],
                    )
            if RUN_AGENTOPS and agentops_session_started:
                agentops.end_session()
            socketio.emit(
                "show_popup",
                {
                    "line1": "Error during analysis process, terminating.",
                    "buttons": ["Ok"],
                },
            )
            wait_for_user_response()
            socketio.emit("analysis_complete", {"index": len(data_store) - 1})
            return "", 204


@app.route("/data_store/<foldername>/<filename>")
def uploaded_file(foldername, filename):
    return send_from_directory("data_store/" + foldername + "/", filename)


@app.route("/get_all_analysis_results")
def get_all_analysis_results():
    return jsonify(data_store)


@app.route("/get_analysis_results/<int:analysis_index>")
def get_analysis_results(analysis_index):
    if 0 <= analysis_index < len(data_store):
        return jsonify(data_store[analysis_index])
    else:
        return jsonify({"error": "Analysis not found"}), 404


@app.route("/start_new_analysis")
def start_new_analysis():
    return jsonify({"message": "Ready for new analysis"})


@app.route("/rename_analysis", methods=["POST"])
def rename_analysis():
    data = request.get_json()
    analysis_index = data["index"]
    new_name = data["new_name"]

    if 0 <= analysis_index < len(data_store):
        data_store[analysis_index]["name"] = new_name
        save_data_store()
        return jsonify({"success": True, "message": "Analysis renamed successfully."})
    else:
        return jsonify({"success": False, "message": "Invalid analysis index."}), 400


@app.route("/update_user_notes", methods=["POST"])
def update_user_notes():
    data = request.get_json(silent=True) or {}
    analysis_index = data.get("index")
    user_notes = data.get("user_notes", "")

    if analysis_index is None or not (0 <= analysis_index < len(data_store)):
        return jsonify({"success": False, "message": "Invalid analysis index."}), 400

    user_notes = (user_notes or "").strip()
    if not user_notes:
        user_notes = DEFAULT_USER_NOTES

    data_store[analysis_index]["user_notes"] = user_notes
    save_data_store()

    return jsonify({"success": True, "message": "User notes updated successfully."})


@app.route("/delete_analysis", methods=["POST"])
def delete_analysis():
    """
    Delete an analysis: remove its folder on disk and its entry in data_store.
    Expects JSON: {"index": <int>}
    """
    data = request.get_json(silent=True) or {}
    analysis_index = data.get("index")

    if analysis_index is None or not (0 <= analysis_index < len(data_store)):
        return jsonify({"success": False, "message": "Invalid analysis index."}), 400

    # delete the folder (if it still exists)
    folder_path = data_store[analysis_index].get("data_folder")
    print(f"Deleting folder: {folder_path}")
    if folder_path and os.path.isdir(folder_path):
        try:
            shutil.rmtree(folder_path)  # recursive delete
        except Exception as exc:
            # If something goes wrong, bail out early so we don’t lose the pointer in data_store while the files are still around.
            return jsonify(
                {
                    "success": False,
                    "message": f"Failed to delete folder: {exc}",
                }
            ), 500

    del data_store[analysis_index]
    save_data_store()

    return jsonify(
        {
            "success": True,
            "message": "Analysis and folder deleted successfully.",
        }
    )


@app.route("/delete_chart", methods=["POST"])
def delete_chart():
    data = request.get_json()
    analysis_index = data["analysis_index"]
    item_index = data["chart_index"]

    if 0 <= analysis_index < len(data_store):
        del data_store[analysis_index]["charts_and_comments"][item_index]
        data_store[analysis_index]["summary"]["update_needed"] = True
        data_store[analysis_index]["report"]["update_needed"] = True
        save_data_store()
        return jsonify({"success": True, "message": "Chart deleted successfully."})
    else:
        return jsonify({"success": False, "message": "Invalid analysis index."}), 400


@app.route("/reorder_analysis_items", methods=["POST"])
def reorder_analysis_items():
    data = request.get_json(silent=True) or {}
    analysis_index = data.get("analysis_index")
    order = data.get("order") or []

    if analysis_index is None or not (0 <= analysis_index < len(data_store)):
        return jsonify({"success": False, "message": "Invalid analysis index."}), 400

    charts = data_store[analysis_index].get("charts_and_comments") or []

    try:
        order_indices = [int(i) for i in order]
    except (TypeError, ValueError):
        return jsonify(
            {"success": False, "message": "Order must be a list of integers."}
        ), 400

    if len(order_indices) != len(charts) or sorted(order_indices) != list(
        range(len(charts))
    ):
        return jsonify(
            {"success": False, "message": "Order must include each item exactly once."}
        ), 400

    data_store[analysis_index]["charts_and_comments"] = [
        charts[i] for i in order_indices
    ]
    save_data_store()

    return jsonify({"success": True, "message": "Analysis items reordered."})


@app.route("/get_code/<int:analysis_index>/<int:chart_index>")
def get_code(analysis_index, chart_index):
    """
    Returns the Python source that created a chart/table.
    """
    try:
        analysis = data_store[analysis_index]
        item = analysis["charts_and_comments"][chart_index]
        fname = item.get("code_filename")
        if not fname:  # no code for this item
            return jsonify({"code": ""}), 404

        path = os.path.join(analysis["data_folder"], fname)
        src = load_file_as_string(path)
        return jsonify({"code": src})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


if __name__ == "__main__":
    load_data_store()
    load_reference_docs()
    prompt_manager = PromptManager(f"prompts/{PROMPT_FILENAME}")
    webbrowser.open_new_tab("http://127.0.0.1:5000")
    socketio.run(app, debug=False)
