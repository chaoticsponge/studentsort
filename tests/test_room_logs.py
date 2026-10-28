from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from exam_access_app import App, ArrangementData, ExamSection, IncidentLogWriter, DEFAULT_DOCX, sections_by_room


class RoomTests(unittest.TestCase):
    def test_missing_room_column_creates_one_combined_file(self):
        data = ArrangementData.__new__(ArrangementData)
        headers = {1: 'First name', 2: 'Surname', 3: 'Access Arrangements',
                   4: 'Subject(s) Applicable'}
        data.years = {
            'Year 11': data._parse_year([headers, {1: 'Alpha', 2: 'Example', 3: 'WP', 4: 'Biology'}]),
            'Year 12': data._parse_year([headers, {1: 'Beta', 2: 'Example', 3: 'SRB', 4: 'Physics'}]),
        }
        data.subject_index = data._build_subject_index()
        app = App.__new__(App)
        app.data = data
        app.selected_exams = lambda: [('Year 11', 'Biology'), ('Year 12', 'Physics')]
        app.update_preview = Mock()
        app.status = Mock()
        app.output_dir = Mock()
        with tempfile.TemporaryDirectory() as folder:
            app.output_dir.get.return_value = folder
            with patch('exam_access_app.messagebox.askyesno', return_value=False):
                app.generate()
            files = list(Path(folder).iterdir())
            self.assertEqual(len(files), 1)
            with ZipFile(files[0]) as archive:
                root = ET.fromstring(archive.read('word/document.xml'))
            text = ' '.join(root.itertext())
            for expected in ['Biology', 'Physics', 'Alpha Example', 'Beta Example']:
                self.assertIn(expected, text)
            self.assertNotIn('Year 11', text)
            self.assertNotIn('Year 12', text)

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
