from django.db.models import ForeignKey
from django.views.generic.base import View
from django.views.generic.list import MultipleObjectMixin

from .response import ExcelResponse


class ExcelMixin(MultipleObjectMixin):
    response_class = ExcelResponse
    header_font = None
    data_font = None
    output_filename = None
    worksheet_name = None
    force_csv = False
    spreadsheet_group_field = None

    def get_header_font(self):
        """
        Return the font to be applied to the header row of the spreadsheet.
        """
        return self.header_font

    def get_data_font(self):
        """
        Return the font to be applied to all data cells in the spreadsheet.
        """
        return self.data_font

    def get_output_filename(self):
        """
        Return the name of the file to be generated, minus the file extension.

        For instance, an output filename of `'excel_data'` would return either
        `'excel_data.xlsx'` or `'excel_data.csv'`
        """
        return self.output_filename or 'excel_data'

    def get_worksheet_name(self):
        """
        Return the name of the worksheet into which the data will be inserted.
        """
        return self.worksheet_name

    def get_force_csv(self):
        """
        Return a boolean for whether or not to force CSV output.
        """
        return self.force_csv

    def get_context_data(self, **kwargs):
        """
        Provide an empty stub since these responses take no context.
        """
        return {}

    def get_excel_queryset(self):
        queryset = self.get_queryset()
        if self.spreadsheet_group_field:
            # check if spreadsheet_group_field str and if it is field of model
            result_dict = dict()
            spreadsheet_field = self.model._meta.get_field(self.spreadsheet_group_field)
            if isinstance(spreadsheet_field, ForeignKey):
                fk_model = spreadsheet_field.remote_field.model
                fk_model_fields_list = [x.name for x in fk_model._meta.fields]


                fk_model_qs = fk_model.objects.all()
                for fk_object in fk_model_qs:

                    fk_field_list = list()
                    for fk_model_field in fk_model_fields_list:
                        fk_field_list.append([fk_model_field, getattr(fk_object, fk_model_field)])
                    fk_field_list.append([])
                    fk_field_list.append([])
                    fk_field_list.append([self.model._meta.verbose_name.title()])
                    fk_field_list.append([])
                    models_field_name_list = [x.name for x in self.model._meta.fields]
                    fk_field_list.append(models_field_name_list)
                    for mod_obje in queryset.filter(**{self.spreadsheet_group_field: fk_object}):
                        fk_field_list.append([str(getattr(mod_obje, x)) for x in models_field_name_list])

                    result_dict[fk_object] = fk_field_list
            else:
                spreadsheet_group_field_list = queryset.order_by().values_list(self.spreadsheet_group_field, flat=True).distinct()
                for spreadsheet_name in spreadsheet_group_field_list:
                    result_dict[spreadsheet_name] = queryset.filter(**{self.spreadsheet_group_field: spreadsheet_name})

            return result_dict
        return queryset


    def render_to_response(self, context, *args, **kwargs):
        return self.response_class(
            self.get_excel_queryset(),
            output_filename=self.get_output_filename(),
            worksheet_name=self.get_worksheet_name(),
            force_csv=self.get_force_csv(),
            header_font=self.get_header_font(),
            data_font=self.get_data_font()
        )


class ExcelView(ExcelMixin, View):
    """
    Return the results of a queryset as an Excel spreadsheet.
    """
    def get(self, request, *args, **kwargs):
        return self.render_to_response({})
