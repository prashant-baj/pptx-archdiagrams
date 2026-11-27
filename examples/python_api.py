"""Use archdiagrams from Python: draw onto any slide of any deck (here a new one)."""
from pptx import Presentation
from pptx.util import Inches

from archdiagrams import Diagram

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# The diagram has its own 12 x 4 in canvas; it is scaled into the box on the slide.
d = Diagram(slide, size=(12, 4), box=(0.6, 1.5, 12.1, 5.0), theme='default')
d.zone('edge', 0, 0, 3.2, 4, 'Edge', color='orange',
       cards=[dict(id='waf', icon='firewall', title='WAF', sub='TLS · rate limit'),
              dict(id='iam', icon='identity', title='Identity', sub='OIDC · MFA')])
d.zone('app', 3.7, 0, 4.0, 4, 'Application', color='blue')
d.card('api', 3.9, 0.6, 3.6, 1.0, 'api', 'Orders API', 'REST · versioned', color='blue')
d.card('q', 3.9, 2.2, 3.6, 1.0, 'queue', 'Job queue', 'at-least-once', color='blue')
d.zone('data', 8.2, 0, 3.8, 4, 'Data', color='green',
       cards=[dict(id='db', icon='database', title='Orders DB', sub='Postgres'),
              dict(id='s3', icon='aws/simple-storage-service-s3', title='Documents', sub='S3')])
d.arrow('waf.r', 'api.l', color='orange')
d.arrow('iam.r', 'api.l@0.75', color='orange', dashed=True)
d.arrow('api.b', 'q.t', color='blue', label='event')
d.arrow('api.r', 'db.l', color='green')
d.arrow('q.r', 's3.l', color='green')
for i, k in enumerate(['waf', 'api', 'q', 'db'], 1):
    d.badge(k, i)
print('warnings:', d.check_overlaps())
prs.save('python_api.pptx')
print('wrote python_api.pptx')
