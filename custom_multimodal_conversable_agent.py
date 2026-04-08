# Portions of this file are adapted from AG2,
# Copyright (c) 2023 - 2025, AG2ai, Inc., AG2ai open-source projects maintainers and core contributors, licensed under the
# Apache License 2.0.
#
# Modified by JK (2025) for subclassing and custom behaviour.


from typing import List, Optional, Union
from autogen.agentchat.contrib.multimodal_conversable_agent import (
    MultimodalConversableAgent,
)
from autogen import Agent, OpenAIWrapper
from autogen.agentchat.contrib.img_utils import message_formatter_pil_to_b64


class CustomMultimodalConversableAgent(MultimodalConversableAgent):
    """Subclass of MultimodalConversableAgent with modified generate_oai_reply (based on AG2 0.7.4).
    Changes:
      -fix tool call issue
      -remove previous images to reduce tokens"""

    def __init__(
        self,
        name: str,
        system_message: Optional[Union[str, List]] = "",
        is_termination_msg: str = None,
        remove_previous_images: bool = True,
        *args,
        **kwargs,
    ):
        self.remove_previous_images = remove_previous_images
        super().__init__(
            name,
            system_message,
            is_termination_msg=is_termination_msg,
            *args,
            **kwargs,
        )
        self.replace_reply_func(
            MultimodalConversableAgent.generate_oai_reply,
            CustomMultimodalConversableAgent.generate_oai_reply,
        )

    def generate_oai_reply(
        self,
        messages: Optional[list[dict]] = None,
        sender: Optional[Agent] = None,
        config: Optional[OpenAIWrapper] = None,
    ) -> tuple[bool, Union[str, dict, None]]:
        """Generate a reply using autogen.oai."""
        client = self.client if config is None else config
        if client is None:
            return False, None
        if messages is None:
            messages = self._oai_messages[sender]

        # EDITS START - Fix tool call issue
        # Check if the last item in the list has 'tool_responses'
        all_messages = []
        for message in messages:
            tool_responses = message.get("tool_responses", [])
            if tool_responses:
                all_messages += tool_responses
                # tool role on the parent message means the content is just concatenation of all of the tool_responses
                if message.get("role") != "tool":
                    all_messages.append(
                        {
                            key: message[key]
                            for key in message
                            if key != "tool_responses"
                        }
                    )
            else:
                all_messages.append(message)

        messages = all_messages
        # END - Fix tool call issue

        # EDITS START - remove previous images
        # Iterate through all dictionaries except the last one
        if self.remove_previous_images:
            for message in messages[:-1]:
                # Check if 'content' exists and is a list
                if "content" in message and isinstance(message["content"], list):
                    # Filter out dictionaries with 'type' == 'image_url'
                    message["content"] = [
                        item
                        for item in message["content"]
                        if not (
                            isinstance(item, dict) and item.get("type") == "image_url"
                        )
                    ]
        # END - remove previous images

        messages_with_b64_img = message_formatter_pil_to_b64(
            self._oai_system_message + messages
        )

        response = client.create(
            context=messages[-1].pop("context", None),
            messages=messages_with_b64_img,
            agent=self.name,
        )

        extracted_response = client.extract_text_or_completion_object(response)[0]
        if not isinstance(extracted_response, str):
            extracted_response = extracted_response.model_dump()
        return True, extracted_response
