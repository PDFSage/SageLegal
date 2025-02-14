from typing import List

class Lawsuit:
    """
    A class representing a lawsuit document that includes a heading, plaintiff, defendant, 
    case information, court name, firm name, ordered body sections, a footer, and ordered exhibits.
    """

    def __init__(
        self,
        heading: str,
        plaintiff: str,
        defendant: str,
        case_information: str,
        court_name: str,
        firm_name: str,
        body_sections: List[str],
        footer: str,
        exhibits: List[str]
    ) -> None:
        """
        Initialize the Lawsuit with the heading, plaintiff, defendant, case information, court name,
        firm name, body sections, footer, and exhibits.

        Parameters:
            heading (str): The title of the lawsuit.
            plaintiff (str): The name of the plaintiff.
            defendant (str): The name of the defendant.
            case_information (str): Additional information about the case (e.g., case number).
            court_name (str): The name of the court handling the case.
            firm_name (str): The name of the firm that files this document.
            body_sections (List[str]): A list of strings representing the ordered body sections.
            footer (str): The footer text of the lawsuit.
            exhibits (List[str]): A list of strings representing the ordered exhibits.
        """
        # Validate required string fields
        for field_name, field_value in [
            ("Heading", heading),
            ("Plaintiff", plaintiff),
            ("Defendant", defendant),
            ("Case information", case_information),
            ("Court name", court_name),
            ("Firm name", firm_name),
            ("Footer", footer)
        ]:
            if not isinstance(field_value, str) or not field_value.strip():
                raise ValueError(f"{field_name} must be a non-empty string.")

        self.heading = heading.strip()
        self.plaintiff = plaintiff.strip()
        self.defendant = defendant.strip()
        self.case_information = case_information.strip()
        self.court_name = court_name.strip()
        self.firm_name = firm_name.strip()
        self.footer = footer.strip()

        # Validate body_sections
        if not isinstance(body_sections, list) or not body_sections:
            raise ValueError("Body sections must be a non-empty list of strings.")
        for idx, section in enumerate(body_sections):
            if not isinstance(section, str) or not section.strip():
                raise ValueError(f"Body section at index {idx} must be a non-empty string.")
        self.body_sections = [section.strip() for section in body_sections]

        # Validate exhibits
        if not isinstance(exhibits, list) or not exhibits:
            raise ValueError("Exhibits must be a non-empty list of strings.")
        for idx, exhibit in enumerate(exhibits):
            if not isinstance(exhibit, str) or not exhibit.strip():
                raise ValueError(f"Exhibit at index {idx} must be a non-empty string.")
        self.exhibits = [exhibit.strip() for exhibit in exhibits]

    def add_body_section(self, section: str) -> None:
        """
        Add a new body section to the lawsuit.

        Parameters:
            section (str): The text for the new body section.
        """
        if not isinstance(section, str) or not section.strip():
            raise ValueError("Section must be a non-empty string.")
        self.body_sections.append(section.strip())

    def add_exhibit(self, exhibit: str) -> None:
        """
        Add a new exhibit to the lawsuit.

        Parameters:
            exhibit (str): The text for the new exhibit.
        """
        if not isinstance(exhibit, str) or not exhibit.strip():
            raise ValueError("Exhibit must be a non-empty string.")
        self.exhibits.append(exhibit.strip())

    def get_full_document(self) -> str:
        """
        Compile the full lawsuit document as a formatted string with header details,
        numbered body sections, and numbered exhibits.

        Returns:
            str: The complete formatted lawsuit document.
        """
        separator = "=" * len(self.heading)
        document_lines = [
            self.heading,
            separator,
            f"Plaintiff: {self.plaintiff}",
            f"Defendant: {self.defendant}",
            f"Firm: {self.firm_name}",
            f"Court: {self.court_name}",
            f"Case Information: {self.case_information}",
            ""
        ]
        document_lines.append("Body Sections:")
        for idx, section in enumerate(self.body_sections, start=1):
            document_lines.append(f"  {idx}. {section}")
        document_lines.append("")
        document_lines.append(self.footer)
        document_lines.append("")
        document_lines.append("Exhibits:")
        for idx, exhibit in enumerate(self.exhibits, start=1):
            document_lines.append(f"  Exhibit {idx}: {exhibit}")
        return "\n".join(document_lines)

    def __str__(self) -> str:
        """
        Return the formatted lawsuit document when the object is printed.

        Returns:
            str: The full formatted lawsuit document.
        """
        return self.get_full_document()


if __name__ == '__main__':
    # Example usage:
    heading = "Lawsuit for Breach of Contract"
    plaintiff = "John Doe"
    defendant = "Acme Corporation"
    case_information = "Case No. 2025-0001"
    court_name = "Superior Court of California"
    firm_name = "Smith & Associates"
    body_sections = [
        "The plaintiff entered into a contract with the defendant on January 1, 2025.",
        "The defendant failed to deliver the services as promised.",
        "The plaintiff incurred significant damages as a result of the breach."
    ]
    footer = "Respectfully submitted, Plaintiff's Attorney."
    exhibits = [
        "Contract Agreement signed on January 1, 2025.",
        "Email correspondence between parties.",
        "Invoice for incurred damages."
    ]

    lawsuit = Lawsuit(
        heading,
        plaintiff,
        defendant,
        case_information,
        court_name,
        firm_name,
        body_sections,
        footer,
        exhibits
    )
    print(lawsuit)