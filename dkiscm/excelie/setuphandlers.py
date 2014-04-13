from collective.grok import gs
from dkiscm.excelie import MessageFactory as _

@gs.importstep(
    name=u'dkiscm.excelie', 
    title=_('dkiscm.excelie import handler'),
    description=_(''))
def setupVarious(context):
    if context.readDataFile('dkiscm.excelie.marker.txt') is None:
        return
    portal = context.getSite()

    # do anything here
