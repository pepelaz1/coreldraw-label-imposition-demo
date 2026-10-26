import unittest
import xml.etree.ElementTree as ET
from planner import plan, svg


class LayoutTests(unittest.TestCase):
    def test_demo_geometry(self):
        p=plan()
        self.assertEqual((p['columns'],p['rows'],p['count']),(4,7,28))
        self.assertAlmostEqual(p['height'],291)
        self.assertAlmostEqual(p['labels'][0]['x'],10)
        self.assertAlmostEqual(p['labels'][1]['x'],60)
    def test_clearance_and_centering(self):
        p=plan(); a,b=p['rails']
        self.assertAlmostEqual(p['labels'][0]['x']-(a['x']+a['width']),1.5)
        self.assertAlmostEqual(b['x']-(p['labels'][3]['x']+50),1.5)
        self.assertAlmostEqual(a['x'],p['web']-b['x']-b['width'])
    def test_compression_after_grid(self):
        a=plan(distortion=1); b=plan(distortion=.97)
        for x,y in zip(a['labels'],b['labels']):
            self.assertAlmostEqual(x['x'],y['x'])
            self.assertAlmostEqual(x['y']*.97,y['y'])
            self.assertAlmostEqual(x['height']*.97,y['height'])
    def test_every_label_inside_web(self):
        for p in (plan(),plan(width=73,height=83,rows=2,columns=2)):
            for x in p['labels']:
                self.assertGreaterEqual(x['x'],0)
                self.assertLessEqual(x['x']+x['width'],p['web'])
                self.assertLessEqual(x['y']+x['height'],p['height']+1e-8)
    def test_oversize_label(self):
        with self.assertRaises(ValueError): plan(width=214)
        with self.assertRaises(ValueError): plan(height=301)
    def test_invalid_numbers(self):
        for kwargs in ({'width':0},{'distortion':0},{'distortion':1.01},
                       {'rows':0},{'columns':5},{'rows':2.2},{'web':float('nan')}):
            with self.assertRaises(ValueError): plan(**kwargs)
    def test_reverse_svg_is_whole_composition(self):
        p=plan(reverse=True); root=ET.fromstring(svg(p))
        group=root.find('{http://www.w3.org/2000/svg}g')
        self.assertIn('scale(-1 1)',group.attrib['transform'])
        self.assertEqual(len(p['labels']),28)
    def test_exact_fit(self):
        p=plan(width=50,height=40,web=207,rapport=280)
        self.assertEqual(p['count'],28)
        self.assertAlmostEqual(p['rails'][0]['x'],0)


if __name__=='__main__': unittest.main()
