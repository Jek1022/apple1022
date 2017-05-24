# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2017-05-24 03:07
from __future__ import unicode_literals

import datetime
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('canvasssheet', '0029_auto_20170523_1406'),
    ]

    operations = [
        migrations.AlterField(
            model_name='csdata',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 5, 24, 11, 7, 21, 664000)),
        ),
        migrations.AlterField(
            model_name='csdetail',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 5, 24, 11, 7, 21, 666000)),
        ),
        migrations.AlterField(
            model_name='csdetail',
            name='postdate',
            field=models.DateTimeField(blank=True, default=datetime.datetime(2017, 5, 24, 11, 7, 21, 666000), null=True),
        ),
        migrations.AlterField(
            model_name='csdetailtemp',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 5, 24, 11, 7, 21, 670000)),
        ),
        migrations.AlterField(
            model_name='csdetailtemp',
            name='postdate',
            field=models.DateTimeField(blank=True, default=datetime.datetime(2017, 5, 24, 11, 7, 21, 670000), null=True),
        ),
        migrations.AlterField(
            model_name='csmain',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 5, 24, 11, 7, 21, 660000)),
        ),
    ]
