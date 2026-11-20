"""Backward-compatible re-export for PDF converters."""

from quickconvert.converters.pdf_converter import register_pdf_converters as register_pdf_image_converters

__all__ = ["register_pdf_image_converters"]
