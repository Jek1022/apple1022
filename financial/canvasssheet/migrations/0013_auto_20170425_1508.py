# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2017-04-25 15:08
from __future__ import unicode_literals

import datetime
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('canvasssheet', '0012_auto_20170424_1431'),
    ]

    operations = [
        migrations.AddField(
            model_name='csdetail',
            name='quantity',
            field=models.IntegerField(default=0),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name='csdetailtemp',
            name='quantity',
            field=models.IntegerField(default=0),
            preserve_default=False,
        ),
        migrations.AlterField(
            model_name='csdata',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 4, 25, 15, 7, 30, 988000)),
        ),
        migrations.AlterField(
            model_name='csdetail',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 4, 25, 15, 7, 30, 990000)),
        ),
        migrations.AlterField(
            model_name='csdetail',
            name='postdate',
            field=models.DateTimeField(blank=True, default=datetime.datetime(2017, 4, 25, 15, 7, 30, 990000), null=True),
        ),
        migrations.AlterField(
            model_name='csdetailtemp',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 4, 25, 15, 7, 31, 2000)),
        ),
        migrations.AlterField(
            model_name='csdetailtemp',
            name='postdate',
            field=models.DateTimeField(blank=True, default=datetime.datetime(2017, 4, 25, 15, 7, 31, 2000), null=True),
        ),
        migrations.AlterField(
            model_name='csmain',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 4, 25, 15, 7, 30, 985000)),
        ),
    ]
