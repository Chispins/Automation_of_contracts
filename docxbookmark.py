from docx.document import Document as _innerdoclass
from docx import Document as _innerdocfn
from docx.oxml.shared import qn
from lxml.etree import Element as El


class Document(_innerdoclass):
    def _bookmark_elements(self, recursive=True):
        if recursive:
            startag = qn('w:bookmarkStart')
            bkms = {}

            def _bookmark_elements_recursive(parent):
                if parent.tag == startag:
                    bookmark_name = parent.attrib.get(qn('w:name'))
                    bookmark_id = str(parent.attrib.get(qn('w:id')))
                    if bookmark_name:
                        bkms[bookmark_id] = {
                            'name': bookmark_name,
                            'current_value': parent.text,
                            'new_value': '',
                        }
                for el in parent:
                    _bookmark_elements_recursive(el)

            for section in self.sections:
                for header in section.header.part.element:
                    _bookmark_elements_recursive(header)

                _bookmark_elements_recursive(self.element)

                for footer in section.footer.part.element:
                    _bookmark_elements_recursive(footer)
            return bkms
        else:
            return self._element.xpath('//' + qn('w:bookmarkStart'))

    def bookmark_names(self):
        """
        Returns a list of bookmarks
        """
        return dict(sorted(self._bookmark_elements().items(), key=lambda x: x[0]))

    def add_bookmark(self, bookmarkname):
        """
        Adds a bookmark with bookmark with name bookmarkname to the end of the file
        """
        el = [el for el in self._element[0] if el.tag.endswith('}p')][-1]
        el.append(El(qn('w:bookmarkStart'), {qn('w:id'): '0', qn('w:name'): bookmarkname}))
        el.append(El(qn('w:bookmarkEnd'), {qn('w:id'): '0'}))

    def __init__(self, innerDocInstance=None):
        super().__init__(Document, None)
        if innerDocInstance is not None and type(innerDocInstance) is _innerdoclass:
            self.__body = innerDocInstance.__body
            self._element = innerDocInstance._element
            self._part = innerDocInstance._part


def DocumentCreate(docx=None):
    """
    Return a |Document| object loaded from *docx*, where *docx* can be
    either a path to a ``.docx`` file (a string) or a file-like object. If
    *docx* is missing or ``None``, the built-in default document "template"
    is loaded.
    """
    return Document(_innerdocfn(docx))