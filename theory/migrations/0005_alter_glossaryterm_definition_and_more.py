# Generated by Django 5.2.1 on 2025-06-02 19:32

import ckeditor.fields
import ckeditor_uploader.fields
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("theory", "0004_statisticsarticle_video_url"),
    ]

    operations = [
        migrations.AlterField(
            model_name="glossaryterm",
            name="definition",
            field=ckeditor.fields.RichTextField(verbose_name="Определение"),
        ),
        migrations.AlterField(
            model_name="statisticsarticle",
            name="content",
            field=ckeditor_uploader.fields.RichTextUploadingField(
                verbose_name="Содержание"
            ),
        ),
        migrations.AlterField(
            model_name="statisticsarticle",
            name="created_at",
            field=models.DateTimeField(auto_now_add=True),
        ),
        migrations.AlterField(
            model_name="statisticsarticle",
            name="title",
            field=models.CharField(max_length=200, verbose_name="Заголовок"),
        ),
    ]
