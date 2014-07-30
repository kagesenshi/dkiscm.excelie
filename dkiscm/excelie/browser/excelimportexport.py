from five import grok
from Products.CMFCore.interfaces import IContentish, ISiteRoot
from zope.schema.interfaces import IVocabularyFactory
from dkiscm.excelie.interfaces import IProductSpecific
import json
from zope.component import getUtility
import tablib
import xlrd
import csv
from StringIO import StringIO
from plone.directives import form
from plone.namedfile.field import NamedFile
from Acquisition import aq_parent
from zope.component.hooks import getSite
from plone.dexterity.utils import createContentInContainer
from Products.statusmessages.interfaces import IStatusMessage

import z3c.form.button

grok.templatedir('templates')

class ExcelImportExport(grok.View):
    grok.context(ISiteRoot)
    grok.name('excelie')
    grok.require('zope2.View')
    grok.template('excelimportexport')


class ExcelExport(grok.View):
    grok.context(ISiteRoot)
    grok.name('excelie-xls')
    grok.layer(IProductSpecific)
    grok.require('cmf.AddPortalContent')

    def render(self):
        output = []
        for brain in self.context.portal_catalog({
                'portal_type': 'dkiscm.jobmatrix.job'
            }):
            obj = brain.getObject()
            output.append(self._get_obj_data(obj))

        headers = []
        for entry in output[0]:
            headers.append(entry[0])

        data = []
        for entry in output:
            data.append([d[1] for d in entry])

        dataset = tablib.Dataset(*data, headers=headers)
        out = dataset.xls
        self.request.response.setHeader('Content-Type', 'application/msword')
        self.request.response.setHeader('Content-Length', len(out))
        self.request.response.setHeader('Content-Transfer-Encoding', 'binary')
        self.request.response.setHeader('Content-Disposition', 
            'attachment; filename=jobmatrix.xls')
        return out

    def _get_obj_data(self, obj):
        job = [
            ('job_code', obj.job_code),
            ('title', obj.Title()),
            ('description', obj.Description()),
            ('industry_cluster', aq_parent(aq_parent(obj)).getId()),
            ('job_grouping', aq_parent(obj).getId()),
            ('education', obj.education),
            ('education_description', obj.education_description),
            ('similar_job_titles', ','.join(obj.similar_job_titles or [])),
            ('professional_certification', ','.join(obj.professional_certification or [])),
            ('job_demand', obj.job_demand),
            ('job_demand_synovate2013', obj.job_demand_synovate2013),
            ('suitable_for_entry', obj.suitable_for_entry)
        ]

        expvocab = getUtility(
            IVocabularyFactory,
            name='dkiscm.jobmatrix.experience'
        )(self.context)

        explevels = []
        for term in expvocab:
            value = term.value
            explevels.append((value, True if value in obj.exp_levels else False))

        salary_range = []        
        for term in expvocab:
            value = term.value
            salary_range.append(
                ('salary_range_%s' % value, obj.salary_range[0][value])
            )

        skills_competency = []
        for idx, skill in enumerate(obj.skills_competency):
            skills_competency.append(
                ('skills_competency_%s_skill' % idx,skill['skill']),
            )
            for term in expvocab:
                value = term.value
                skills_competency.append(
                    ('skills_competency_%s_%s' % (idx, value), skill[value])
                )
                skills_competency.append(
                    ('skills_competency_%s_%s_required' % (idx, value), 
                        skill['%s_required' % value])
                )

        for idx in range(len(obj.skills_competency), 9):
            skills_competency.append(
                ('skills_competency_%s_skill' % idx, None),
            )
            for term in expvocab:
                value = term.value
                skills_competency.append(
                    ('skills_competency_%s_%s' % (idx, value), None)
                )
                skills_competency.append(
                    ('skills_competency_%s_%s_required' % (idx, value),
                        False)
                )


        softskills_competency = []
        for idx, skill in enumerate(obj.softskills_competency):
            softskills_competency.append(
                ('softskills_competency_%s_skill' % idx,skill['skill']),
            )
            for term in expvocab:
                value = term.value
                softskills_competency.append(
                    ('softskills_competency_%s_%s' % (idx, value), skill[value])
                )
                softskills_competency.append(
                    ('softskills_competency_%s_%s_weight' % (idx, value), 
                        skill['%s_weight' % value])
                )


        for idx in range(len(obj.softskills_competency), 6):
            softskills_competency.append(
                ('softskills_competency_%s_skill' % idx, None),
            )
            for term in expvocab:
                value = term.value
                softskills_competency.append(
                    ('softskills_competency_%s_%s' % (idx, value), None)
                )
                softskills_competency.append(
                    ('softskills_competency_%s_%s_weight' % (idx, value), 1)
                )

        return job + explevels + salary_range + skills_competency + softskills_competency


class IUploadFormSchema(form.Schema):
    import_file = NamedFile(title=u'Upload Excel')

class UploadForm(form.SchemaForm):
    name = u"Import JobMatrix from Excel"
    schema = IUploadFormSchema
    ignoreContext = True
    grok.context(ISiteRoot)
    grok.name('excelie-import')
    grok.require('cmf.AddPortalContent')

    _field_titles = ['job_code',
        'title',
        'description',
        'industry_cluster',
        'job_grouping',
        'education',
        'education_description',
        'similar_job_titles',
        'professional_certification',
        'job_demand',
        'job_demand_synovate2013',
        'suitable_for_entry',
        'entry',
        'intermediate',
        'senior',
        'advanced',
        'master',
        'salary_range_entry',
        'salary_range_intermediate',
        'salary_range_senior',
        'salary_range_advanced',
        'salary_range_master',
        'skills_competency_0_skill',
        'skills_competency_0_entry',
        'skills_competency_0_entry_required',
        'skills_competency_0_intermediate',
        'skills_competency_0_intermediate_required',
        'skills_competency_0_senior',
        'skills_competency_0_senior_required',
        'skills_competency_0_advanced',
        'skills_competency_0_advanced_required',
        'skills_competency_0_master',
        'skills_competency_0_master_required',
        'skills_competency_1_skill',
        'skills_competency_1_entry',
        'skills_competency_1_entry_required',
        'skills_competency_1_intermediate',
        'skills_competency_1_intermediate_required',
        'skills_competency_1_senior',
        'skills_competency_1_senior_required',
        'skills_competency_1_advanced',
        'skills_competency_1_advanced_required',
        'skills_competency_1_master',
        'skills_competency_1_master_required',
        'skills_competency_2_skill',
        'skills_competency_2_entry',
        'skills_competency_2_entry_required',
        'skills_competency_2_intermediate',
        'skills_competency_2_intermediate_required',
        'skills_competency_2_senior',
        'skills_competency_2_senior_required',
        'skills_competency_2_advanced',
        'skills_competency_2_advanced_required',
        'skills_competency_2_master',
        'skills_competency_2_master_required',
        'skills_competency_3_skill',
        'skills_competency_3_entry',
        'skills_competency_3_entry_required',
        'skills_competency_3_intermediate',
        'skills_competency_3_intermediate_required',
        'skills_competency_3_senior',
        'skills_competency_3_senior_required',
        'skills_competency_3_advanced',
        'skills_competency_3_advanced_required',
        'skills_competency_3_master',
        'skills_competency_3_master_required',
        'skills_competency_4_skill',
        'skills_competency_4_entry',
        'skills_competency_4_entry_required',
        'skills_competency_4_intermediate',
        'skills_competency_4_intermediate_required',
        'skills_competency_4_senior',
        'skills_competency_4_senior_required',
        'skills_competency_4_advanced',
        'skills_competency_4_advanced_required',
        'skills_competency_4_master',
        'skills_competency_4_master_required',
        'skills_competency_5_skill',
        'skills_competency_5_entry',
        'skills_competency_5_entry_required',
        'skills_competency_5_intermediate',
        'skills_competency_5_intermediate_required',
        'skills_competency_5_senior',
        'skills_competency_5_senior_required',
        'skills_competency_5_advanced',
        'skills_competency_5_advanced_required',
        'skills_competency_5_master',
        'skills_competency_5_master_required',
        'skills_competency_6_skill',
        'skills_competency_6_entry',
        'skills_competency_6_entry_required',
        'skills_competency_6_intermediate',
        'skills_competency_6_intermediate_required',
        'skills_competency_6_senior',
        'skills_competency_6_senior_required',
        'skills_competency_6_advanced',
        'skills_competency_6_advanced_required',
        'skills_competency_6_master',
        'skills_competency_6_master_required',
        'skills_competency_7_skill',
        'skills_competency_7_entry',
        'skills_competency_7_entry_required',
        'skills_competency_7_intermediate',
        'skills_competency_7_intermediate_required',
        'skills_competency_7_senior',
        'skills_competency_7_senior_required',
        'skills_competency_7_advanced',
        'skills_competency_7_advanced_required',
        'skills_competency_7_master',
        'skills_competency_7_master_required',
        'skills_competency_8_skill',
        'skills_competency_8_entry',
        'skills_competency_8_entry_required',
        'skills_competency_8_intermediate',
        'skills_competency_8_intermediate_required',
        'skills_competency_8_senior',
        'skills_competency_8_senior_required',
        'skills_competency_8_advanced',
        'skills_competency_8_advanced_required',
        'skills_competency_8_master',
        'skills_competency_8_master_required',
        'softskills_competency_0_skill',
        'softskills_competency_0_entry',
        'softskills_competency_0_entry_weight',
        'softskills_competency_0_intermediate',
        'softskills_competency_0_intermediate_weight',
        'softskills_competency_0_senior',
        'softskills_competency_0_senior_weight',
        'softskills_competency_0_advanced',
        'softskills_competency_0_advanced_weight',
        'softskills_competency_0_master',
        'softskills_competency_0_master_weight',
        'softskills_competency_1_skill',
        'softskills_competency_1_entry',
        'softskills_competency_1_entry_weight',
        'softskills_competency_1_intermediate',
        'softskills_competency_1_intermediate_weight',
        'softskills_competency_1_senior',
        'softskills_competency_1_senior_weight',
        'softskills_competency_1_advanced',
        'softskills_competency_1_advanced_weight',
        'softskills_competency_1_master',
        'softskills_competency_1_master_weight',
        'softskills_competency_2_skill',
        'softskills_competency_2_entry',
        'softskills_competency_2_entry_weight',
        'softskills_competency_2_intermediate',
        'softskills_competency_2_intermediate_weight',
        'softskills_competency_2_senior',
        'softskills_competency_2_senior_weight',
        'softskills_competency_2_advanced',
        'softskills_competency_2_advanced_weight',
        'softskills_competency_2_master',
        'softskills_competency_2_master_weight',
        'softskills_competency_3_skill',
        'softskills_competency_3_entry',
        'softskills_competency_3_entry_weight',
        'softskills_competency_3_intermediate',
        'softskills_competency_3_intermediate_weight',
        'softskills_competency_3_senior',
        'softskills_competency_3_senior_weight',
        'softskills_competency_3_advanced',
        'softskills_competency_3_advanced_weight',
        'softskills_competency_3_master',
        'softskills_competency_3_master_weight',
        'softskills_competency_4_skill',
        'softskills_competency_4_entry',
        'softskills_competency_4_entry_weight',
        'softskills_competency_4_intermediate',
        'softskills_competency_4_intermediate_weight',
        'softskills_competency_4_senior',
        'softskills_competency_4_senior_weight',
        'softskills_competency_4_advanced',
        'softskills_competency_4_advanced_weight',
        'softskills_competency_4_master',
        'softskills_competency_4_master_weight',
        'softskills_competency_5_skill',
        'softskills_competency_5_entry',
        'softskills_competency_5_entry_weight',
        'softskills_competency_5_intermediate',
        'softskills_competency_5_intermediate_weight',
        'softskills_competency_5_senior',
        'softskills_competency_5_senior_weight',
        'softskills_competency_5_advanced',
        'softskills_competency_5_advanced_weight',
        'softskills_competency_5_master',
        'softskills_competency_5_master_weight']
    

    @z3c.form.button.buttonAndHandler(u"Import", name='import')
    def import_content(self, action):
        formdata, errors = self.extractData()
        if errors:
            self.status = self.formErrorsMessage
            return

        self._import(formdata['import_file'].data)


    def _import(self, xls):
        data = self._to_json(xls)

        counter = {'update': 0, 'create': 0}
        for entry in data:
            self._create(entry, counter)

        IStatusMessage(self.request).addStatusMessage(
            u"%(update)s items updated, %(create)s items created" % counter
        )

    def _create(self, data, counter):
        brains = self.context.portal_catalog({'getId': data['job_code'].lower()})        
        if brains:
            self._update(brains[0].getObject(), data)
            counter['update'] += 1
            return

        container = self._find_container(
                data['industry_cluster'],
                data['job_grouping'],
        )

        obj = createContentInContainer(container, 'dkiscm.jobmatrix.job',
                                        title=data['title'],
                                        job_code=data['job_code'])
        
        self._update(obj, data)
        counter['create'] += 1

    def _update(self, obj, data):
        data = self._cook_data(data)

        obj.setTitle(data['title'])
        obj.setDescription(data['description'])

        for k in ['job_code', 'education', 'education_description',
                  'similar_job_titles', 'exp_levels', 
                  'professional_certification',
                  'job_demand',
                  'job_demand_synovate2013',
                  'suitable_for_entry']:
            setattr(obj, k, data[k])

        for k in ['salary_range', 'skills_competency',
                'softskills_competency']:
            setattr(obj, k, data[k])

    def _cook_data(self, data):
        for i in ['similar_job_titles', 'professional_certification']:
            data[i] = [v.strip() for v in data[i].split(',') if v.strip()]

        for intkey in ['job_demand','job_demand_synovate2013']:
            data[intkey] = int(data[intkey])

        for boolkey in ['suitable_for_entry']:
            data[boolkey] = bool(int(data[boolkey]))

        expvocab = getUtility(
            IVocabularyFactory,
            name='dkiscm.jobmatrix.experience'
        )(self.context)

        key = 'exp_levels'
        data[key] = []

        for i in [term.value for term in expvocab]:
            if bool(int(data[i])):
               data[key].append(i)

        key = 'salary_range'
        data[key] = []
        edata = {}
        for exp in [term.value for term in expvocab]:
            edata[exp] = data['%s_%s' % (key, exp)]
        data[key].append(edata)

        key = 'skills_competency'
        data[key] = []
        for i in xrange(9):
            edata = {'skill': data['%s_%s_skill' % (key, i)]}
            for exp in [term.value for term in expvocab]:
                edata[exp] = data['%s_%s_%s' % (key, i, exp)]
                edata['%s_required' % exp] = bool(data[
                    '%s_%s_%s_required' % (key, i, exp)
                ])
            data[key].append(edata)

        key = 'softskills_competency'
        data[key] = []
        for i in xrange(6):
            edata = {'skill': data['%s_%s_skill' % (key, i)]}
            for exp in [term.value for term in expvocab]:
                edata[exp] = data['%s_%s_%s' % (key, i, exp)]
                edata['%s_weight' % exp] = int(data[
                    '%s_%s_%s_weight' % (key, i, exp)
                ])
            data[key].append(edata)

        return data            

    def _find_container(self, industry_cluster, job_grouping):
        site = getSite()
        if not 'cluster' in site.keys():
            site.invokeFactory(type_name='Folder', id='cluster')
            obj = site['cluster']
            obj.setTitle('Clusters')
            obj.reindexObject()
        repo = site['cluster']
        if not repo.has_key(industry_cluster):
            raise Exception(
                ('Unable to locate industry cluster %s,'
                 'please create it first') % industry_cluster
            )
        cluster = repo[industry_cluster]

        if not cluster.has_key(job_grouping):
            raise Exception(
            ('Unable to locate job group %s, for cluster %s'
                 'please create it first') % (job_grouping, industry_cluster)
            )
        container = cluster[job_grouping]
        return container


    def _to_json(self, xls):
        wb = xlrd.open_workbook('jobmatrix.xls', file_contents=xls)
        sh = wb.sheet_by_name('Tablib Dataset')
        data = []
        for rownum in xrange(1, sh.nrows):
            rowdata = {}
            values = sh.row_values(rownum)
            for i in xrange(len(self._field_titles)):
                rowdata[self._field_titles[i]] = values[i]
            data.append(rowdata)
        return data
