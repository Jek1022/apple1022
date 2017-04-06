# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2017-04-06 12:00
from __future__ import unicode_literals

import datetime
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('purchaserequisitionform', '0009_auto_20170406_1142'),
    ]

    operations = [
        migrations.AlterModelOptions(
            name='prfmain',
            options={'ordering': ['-pk'], 'permissions': (('view_purchaserequisitionform', 'Can view purchaserequisitionform'), ('view_assignprf', 'Can view only assigned prf'), ('view_allassignprf', 'Can view all prf'), ('can_approveprf', 'Can approve prf'), ('can_disapproveprf', 'Can disapprove prf'))},
        ),
        migrations.AlterField(
            model_name='prfdetail',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 4, 6, 12, 0, 18, 522000)),
        ),
        migrations.AlterField(
            model_name='prfdetail',
            name='postdate',
            field=models.DateTimeField(blank=True, default=datetime.datetime(2017, 4, 6, 12, 0, 18, 522000), null=True),
        ),
        migrations.AlterField(
            model_name='prfdetailtemp',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 4, 6, 12, 0, 18, 524000)),
        ),
        migrations.AlterField(
            model_name='prfmain',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 4, 6, 12, 0, 18, 519000)),
        ),
    ]
