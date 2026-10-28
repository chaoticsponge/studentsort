from pathlib import Path
import tempfile
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from exam_access_app import ArrangementData, ExamSection, IncidentLogWriter, DEFAULT_DOCX, sections_by_room


class RoomTests(unittest.TestCase):
    def test_parse_group_and_generate(self):
        parser = ArrangementData.__new__(ArrangementData)
        rows = [{1: 'First name', 2: 'Surname', 3: 'Access Arrangements',
                 4: 'Subject(s) Applicable', 5: 'Room'}]
        for name, subject, room in [('Alpha', 'Biology', 'Dance studio'),
                                    ('Beta', 'Physics', 'DANCE STUDIO'),
                                    ('Gamma', 'Physics', 'Board room'),
                                    ('Delta', 'Physics', ' ALTERNATIVE ROOM '),
                                    ('Epsilon', 'Physics', '')]:
            rows.append({1: name, 2: 'Example', 3: 'WP', 4: subject, 5: room})
        students = parser._parse_year(rows)
        sections = tuple(ExamSection(subject, tuple(s for s in students if subject in s.subjects))
                         for subject in ['Biology', 'Physics'])
        rooms = sections_by_room(sections)
        self.assertEqual(set(rooms), {'Dance studio', 'Board room', 'Room not specified'})
        with tempfile.TemporaryDirectory() as folder:
            for room, groups in rooms.items():
                path = Path(folder) / (room + '.docx')
                IncidentLogWriter(DEFAULT_DOCX).write_sections(path, groups)
                with ZipFile(path) as archive:
                    root = ET.fromstring(archive.read('word/document.xml'))
                text = ' '.join(root.itertext())
                self.assertNotIn('Delta Example', text)
                if room == 'Dance studio':
                    self.assertIn('Alpha Example', text)
                    self.assertIn('Beta Example', text)
                    self.assertNotIn('Gamma Example', text)
            self.assertEqual(len(list(Path(folder).glob('*.docx'))), 3)


if __name__ == '__main__':
    unittest.main()
