# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2018-02-01 06:30
from __future__ import unicode_literals

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('processing_jv', '0004_auto_20180201_0931'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='logs_jvmain',
            name='year',
        ),
        migrations.RemoveField(
            model_name='logs_jvmain',
            name='year2',
        ),
        migrations.AddField(
            model_name='logs_jvdetail',
            name='status',
            field=models.CharField(blank=True, max_length=500, null=True),
        ),
    ]
