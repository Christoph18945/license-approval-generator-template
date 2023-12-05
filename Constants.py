#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Constant variables with string values for license approval generation.


"""
from typing import *

class Constants:
    """
    Class contains getter methods to retrieve all
    constants for the license approval composition.

    The costants are the necessary properties text for the tables
    in the DOCX file.
    """

    def __init__(self):
        self.HEADER_TEXT: str = ""
        self.TITLE_MAIN_SECTION: str = ""
        self.SENDER_FORM: str = ""
        self.SUBMISSION_TEXT_PROPERTY: str = ""
        self.SUBMISSION_TEXT_NAME: str = ""
        self.SUBMISSION_TO: str = ""
        self.APPT_OR_REJ_TEXT: str = ""
        self.SUBM_DATE: str = ""
        self.DATE_APPR_TEXT: str = ""
        self.THIRD_PARTY_NAME_PROP: str = ""
        self.VERSION_YEAR_PROP: str = ""
        self.UPDATE_PROP: str = ""
        self.SOFTW_DESC_PROP: str = ""
        self.LINK_PROPERTY_PROP: str = ""
        self.LICENSE_PROP: str = ""
        self.LINK_LIC_PROP: str = ""
        self.PROD_PROP: str = ""
        self.TIME_VER_PROP: str = ""

    # getter methods
    # header section
    def get_header_text(self) -> str:
        self.HEADER_TEXT = "INTERNAL USE ONLY"
        return self.HEADER_TEXT

    def get_title_main_section(self) -> str:
        self.TITLE_MAIN_SECTION = "THIRD PARTY SOFTWARE LICENSE APPROVAL FORM"    
        return self.TITLE_MAIN_SECTION

    # meta info section
    def get_sender_form(self) -> str:
        self.SENDER_FORM = "From: Christoph Hartleb (Dev)"
        return self.SENDER_FORM

    def get_submission_text_property(self) -> str:
        self.SUBMISSION_TEXT_PROPERTY = "Submitted to Legal by:"
        return self.SUBMISSION_TEXT_PROPERTY

    def get_submission_text_name(self) -> str:
        self.SUBMISSION_TEXT_NAME = "Christoph Hartleb"
        return self.SUBMISSION_TEXT_NAME

    def get_submission_to(self) -> str:
        self.SUBMISSION_TO = "To: Lawyer"
        return self.SUBMISSION_TO

    def get_appt_or_rej_text(self) -> str:
        self.APPT_OR_REJ_TEXT = "Approved/Rejected by Legal:"
        return self.APPT_OR_REJ_TEXT

    def get_sub_date(self) -> str:
        self.SUBM_DATE = "Submission Date: "
        return self.SUBM_DATE

    def get_date_appr_text(self) -> str:
        self.DATE_APPR_TEXT = "Date Approved:"
        return self.DATE_APPR_TEXT

    def get_date_format(self) -> str:
        self.DATE_FORMAT = "MM/DD/YYYY"
        return self.DATE_FORMAT

    # main section
    def get_third_party_name_prop(self) -> str:
        self.THIRD_PARTY_NAME_PROP = "Name of third party software:"
        return self.THIRD_PARTY_NAME_PROP

    def get_version_year_prop(self) -> str:
        self.VERSION_YEAR_PROP = "Version number or year:"
        return self.VERSION_YEAR_PROP

    def get_update_prop(self) -> str:
        self.UPDATE_PROP = """Is this a version update of 
previously approved software? If 
Yes, reason for update?"""
        return self.UPDATE_PROP

    def get_softw_desc_prop(self) -> str:
        self.SOFTW_DESC_PROP = "General description of software:"
        return self.SOFTW_DESC_PROP

    def get_link_property_prop(self) -> str:
        self.LINK_PROPERTY_PROP = "Link to software homepage:"
        return self.LINK_PROPERTY_PROP

    def get_license_prop(self) -> str:
        self.LICENSE_PROP = "License type (e.g. MIT, BSD, GPL)"
        return self.LICENSE_PROP

    def get_link_lic_prop(self) -> str:
        self.LINK_LIC_PROP = "Link to website showing license:"
        return self.LINK_LIC_PROP

    def get_prod_prop(self) -> str:
        self.ALTOVA_PROD_PROP = """Company products that will
introduce license?"""
        return self.PROD_PROP

    def get_affected_products(self) -> str:
        self.AFFECTED_PRODUCTS = """All products that support the packages."""
        return self.AFFECTED_PRODUCTS

    def get_time_ver_prop(self) -> str:
        self.TIME_VER_PROP = "Approximate time/version?"
        return self.TIME_VER_PROP
