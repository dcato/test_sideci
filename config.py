#!/usr/bin/env python2
# vim:fileencoding=UTF-8:ts=4:sw=4:sta:et:sts=4:ai

__license__   = 'GPL v3'
__copyright__ = '2016, D.Cato'
__docformat__ = 'restructuredtext en'

import sys, copy, json, traceback
from xml.sax.saxutils import escape

from PyQt5.Qt import (QObject, QApplication, QWidget, QIcon, QTextCursor,
        pyqtSignal, QDialog, QDialogButtonBox, QVBoxLayout)


from calibre import isbytestring, force_unicode, as_unicode

from calibre.utils.config import JSONConfig
from calibre.utils.logging import GUILog, Log, HTMLStream, ANSIStream

from calibre.gui2 import (error_dialog, question_dialog, info_dialog,
        choose_files, choose_save_file)
from calibre.gui2.widgets import PythonHighlighter

from calibre_plugins.exec_macro.config_ui import Ui_Form


# This is where all preferences for this plugin will be stored
# Remember that this name (i.e. plugins/interface_demo) is also
# in a global namespace, so make it as unique as possible.
# You should always prefix your config file name with plugins/,
# so as to ensure you dont accidentally clobber a calibre config file
prefs = JSONConfig('plugins/ExecMacro')

# Set default custom columns
prefs.defaults['current_macro'] = 'Sample Dialog'
prefs.defaults['macros'] = {'Sample Dialog':
{
'name': 'Sample Dialog',
'documentation': 'Macro to display simple dialog.',
'program': '''\
from calibre.gui2 import info_dialog
result = info_dialog(self.gui, 'Greetings',
    'Hello World. Welcome to ExecMacro Plugin.', show=True)
''',
'filename': '',
},
}


help_text = _('''
<p>Here you can add and remove macros executed from Exec Macro plugin. A
macro is written in python. It takes information from the
book, processes it in some way, then returns a string result. Functions
defined here are usable in templates in the same way that builtin
functions are usable. The function must be named <b>evaluate</b>, and
must have the signature shown below.</p>
<p><code>evaluate(self, formatter, kwargs, mi, locals, your parameters)
&rarr; returning a unicode string</code></p>
<p>The parameters of the evaluate function are:
<ul>
<li><b>formatter</b>: the instance of the formatter being used to
evaluate the current template. You can use this to do recursive
template evaluation.</li>
<li><b>kwargs</b>: a dictionary of metadata. Field values are in this
dictionary.
<li><b>mi</b>: a Metadata instance. Used to get field information.
This parameter can be None in some cases, such as when evaluating
non-book templates.</li>
<li><b>locals</b>: the local variables assigned to by the current
template program.</li>
<li><b>your parameters</b>: You must supply one or more formal
parameters. The number must match the arg count box, unless arg count is
-1 (variable number or arguments), in which case the last argument must
be *args. At least one argument is required, and is usually the value of
the field being operated upon. Note that when writing in basic template
mode, the user does not provide this first argument. Instead it is
supplied by the formatter.</li>
</ul></p>
<p>
The following example function checks the value of the field. If the
field is not empty, the field's value is returned, otherwise the value
EMPTY is returned.
<pre>
name: my_ifempty
arg count: 1
doc: my_ifempty(val) -- return val if it is not empty, otherwise the string 'EMPTY'
program code:
def evaluate(self, formatter, kwargs, mi, locals, val):
    if val:
        return val
    else:
        return 'EMPTY'</pre>
</p>
''')

config_dialog = None
config_widget = None

def show_config_dialog(ia):
    global config_dialog
    global config_widget

    if not config_dialog:
        config_dialog = QDialog(ia.gui)

        config_widget = ConfigWidget(ia)

        button_box = QDialogButtonBox(QDialogButtonBox.Close)
        button_box.accepted.connect(hide_config_dialog)
        button_box.rejected.connect(hide_config_dialog)

        v = QVBoxLayout(config_dialog)
        v.addWidget(config_widget)
        v.addWidget(button_box)

        geom = prefs.get('config_dialog_geometry')
        if geom:
            config_dialog.restoreGeometry(geom)
        else:
            config_dialog.resize(config_dialog.sizeHint())

        config_dialog.setModal(False)

    config_dialog.show()
    config_dialog.raise_()
    config_dialog.activateWindow()
    config_widget.textBrowser.setHtml(help_text)
    config_widget.scroll_function_names_box(prefs['current_macro'])

def hide_config_dialog():
    prefs['config_dialog_geometry'] = bytearray(config_dialog.saveGeometry())
    config_dialog.hide()


class ConfigWidget(QWidget, Ui_Form):

    def __init__(self, ia):
        QWidget.__init__(self)
        self.setupUi(self)
        self.ia = ia

        self.initialize()


    def initialize(self):
        self.function_name.currentIndexChanged[str].connect(self.function_index_changed)
        self.function_name.editTextChanged.connect(self.function_name_edited)
        self.documentation.textChanged.connect(self.enable_save_button)
        self.program.textChanged.connect(self.enable_save_button)
        self.macrofile.textChanged.connect(self.enable_save_button)
        self.fromfile_checkbox.stateChanged.connect(self.fromfile_checkbox_clicked)

        self.delete_button.clicked.connect(self.delete_button_clicked)
        self.save_button.clicked.connect(self.save_button_clicked)

        self.execute_button.clicked.connect(self.execute_button_clicked)
        self.import_button.clicked.connect(self.import_button_clicked)
        self.export_button.clicked.connect(self.export_button_clicked)
        self.save_file_button.clicked.connect(self.save_file_button_clicked)
        self.load_file_button.clicked.connect(self.load_file_button_clicked)
        self.filebrowse_button.clicked.connect(self.filebrowse_button_clicked)
        self.filebrowse_button.setIcon(QIcon(I('document_open.png')))

        self.save_button.setEnabled(False)
        self.delete_button.setEnabled(False)
        self.program.setTabStopWidth(20)
        self.highlighter = PythonHighlighter(self.program.document())

        self.build_function_names_box()

    def execute_button_clicked(self):
        try:
            log = GUILog()
            log.outputs.append(QtLogStream())
            log.outputs[1].html_log.connect(self.log_outputted)
            log.outputs.append(ANSIStream())

            if self.fromfile_checkbox.isChecked():
                self.load_file_button_clicked()

            program = self.program.toPlainText()
            encoding = self.ia.get_encoding(program)
            if encoding:
                program = program.encode(encoding)

            self.textBrowser.clear()

            self.ia.execute(program, log)
        except:
            log.exception('Failed to execute macro:')

            lineno = 0
            for tb in traceback.extract_tb(sys.exc_info()[2]):
                if tb[0] == '<string>':
                    lineno = tb[1]
            if 0 < lineno:
                self.program.go_to_line(lineno)


    def import_button_clicked(self):
        filenames = choose_files(self, 'Exec Macro: Import macros'
            , _('Import macros'), select_only_single_file=True)
        if not filenames:
            return

        try:
            with open(filenames[0], 'r') as fd:
                new_macros = json.load(fd)
            overwrite = sorted(new_macros.viewkeys() & prefs['macros'].viewkeys())
            if not question_dialog(self, _('Some macros will be overwitten'),
                    _('Imported file contains already exists macro,'
                        ' so some macros will be overwitten.'
                        ' Do you want to proceed?'), det_msg='\n'.join(overwrite)):
                return

            prefs['macros'].update(new_macros)
            self.build_function_names_box()
            self.scroll_function_names_box(self.function_name.currentText())
            self.ia.rebuild_menu()
            prefs.commit()

            info_dialog(self, _('Macros imported'),
                _('%d macros imported.') % len(new_macros), show=True)
        except:
            error_dialog(self, _('Failed to import macros'), _(
                'Failed to import macros, click "Show details" for more information.'),
                 det_msg=traceback.format_exc(), show=True)


    def export_button_clicked(self):
        filename = choose_save_file(self, 'Exec Macro: Export macros',
            _('Export macros'))
        if not filename:
            return
        try:

            with open(filename, 'w') as fd:
                json.dump(prefs['macros'], fd, sort_keys=True, indent=4)

            info_dialog(self, _('Macros exported'),
                _('%d macros exported.') % len(prefs['macros']), show=True)
        except:
            error_dialog(self, _('Failed to export macros'), _(
                'Failed to export macros, click "Show details" for more information.'),
                 det_msg=traceback.format_exc(), show=True)


    def save_file_button_clicked(self):
        filename = self.macrofile.text()
        if not filename:
            filename = choose_save_file(self, 'Exec Macro: Save macro',
                _('Save macro'))
            self.macrofile.setText(filename)
        if not filename:
            return

        try:
            program = self.program.toPlainText()
            encoding = self.ia.get_encoding(program)
            if encoding:
                program = program.encode(encoding)
            with open(filename, 'w') as fd:
                fd.write(program)

            info_dialog(self, _('Macro saved to file'),
                _("Current macro's program code saved to a file."), show=True)

        except:
            error_dialog(self, _('Failed to save macro'), _(
                'Failed to save macro, click "Show details" for more information.'),
                 det_msg=traceback.format_exc(), show=True)

    def load_file_button_clicked(self):
        filename = self.macrofile.text()
        if not filename:
            filename = self.filebrowse_button_clicked()
        if not filename:
            return

        try:
            with open(self.macrofile.text(), 'rU') as fd:
                program = fd.read()
                encoding = self.ia.get_encoding(program)
                if encoding:
                    program = program.decode(encoding)

            self.program.setPlainText(program)

            info_dialog(self, _('Macro loaded from file'),
                _("Current macro's program code loaded from a file."), show=True)
        except:
            error_dialog(self, _('Failed to load macro'), _(
                'Failed to load macro, click "Show details" for more information.'),
                 det_msg=traceback.format_exc(), show=True)

    def filebrowse_button_clicked(self):
        filenames = choose_files(self, 'Exec Macro: load macro', _('Load macro'),
            select_only_single_file=True, default_dir=self.macrofile.text())
        if filenames:
            self.macrofile.setText(filenames[0])
            return filenames[0]

    def fromfile_checkbox_clicked(self, i):
        self.save_button.setEnabled(True)
        self.program.setReadOnly(self.fromfile_checkbox.isChecked())

    def enable_save_button(self):
        self.save_button.setEnabled(True)

    def delete_button_clicked(self):
        name = self.function_name.currentText()
        if name in prefs['macros']:
            del prefs['macros'][name]
            prefs.commit()
            self.save_button.setEnabled(True)
            self.delete_button.setEnabled(False)
            self.build_function_names_box(set_to=name)
            self.ia.rebuild_menu()
        else:
            error_dialog(self, _('Exec Macro'),
                         _('Macro not defined'), show=True)

    def build_function_names_box(self, set_to=''):
        self.function_name.blockSignals(True)
        func_names = sorted(prefs['macros'])
        self.function_name.clear()
        self.function_name.addItem('')
        self.function_name.addItems(func_names)
        self.function_name.setCurrentIndex(0)
        if set_to:
            self.function_name.setEditText(set_to)
            self.save_button.setEnabled(True)
        self.function_name.blockSignals(False)


    def scroll_function_names_box(self, scroll_to):
        idx = self.function_name.findText(scroll_to)
        if idx >= 0:
            self.function_name.setCurrentIndex(idx)
            self.delete_button.setEnabled(True)


    def save_button_clicked(self):
        name = self.function_name.currentText()

        prefs['macros'][name] = {
            'name': name,
            'macrofile': self.macrofile.text(),
            'execfromfile': self.fromfile_checkbox.isChecked(),
            'documentation': self.documentation.toPlainText(),
            'program': self.program.toPlainText(),
        }
        self.build_function_names_box()
        self.scroll_function_names_box(name)
        self.ia.rebuild_menu()
        prefs.commit()


    def function_name_edited(self, txt):
        if not txt:
            self.delete_button.setEnabled(False)
            self.save_button.setEnabled(False)
        elif prefs['macros'].get(txt):
            self.delete_button.setEnabled(True)
            self.save_button.setEnabled(False)
        else:
            self.delete_button.setEnabled(False)
            self.save_button.setEnabled(True)

    def function_index_changed(self, txt):
        if not txt:
            self.program.clear()
            self.macrofile.setText('')
            self.fromfile_checkbox.setChecked(False)
            self.documentation.clear()
            self.delete_button.setEnabled(False)
            self.save_button.setEnabled(False)
        else:
            func = prefs['macros'][txt]
            self.program.setPlainText(func.get('program'))
            self.macrofile.setText(func.get('macrofile'))
            self.fromfile_checkbox.setChecked(func.get('execfromfile', False))
            self.documentation.setText(func.get('documentation'))
            self.delete_button.setEnabled(True)
            self.save_button.setEnabled(False)

    def log_outputted(self, txt):
        self.textBrowser.moveCursor(QTextCursor.End)
        text = (u'<pre style="margin-top:0px; margin-bottom:0px;">'
                + txt + u'</pre>')
        self.textBrowser.insertHtml(text)
        self.textBrowser.moveCursor(QTextCursor.End)
        self.textBrowser.repaint()
        QApplication.processEvents()

    def refresh_gui(self, gui):
        pass


class QtLogStream(HTMLStream, QObject):

    html_log = pyqtSignal(unicode)
    plain_text_log = pyqtSignal(unicode)


    def __init__(self):
        HTMLStream.__init__(self)
        QObject.__init__(self)

    def prints(self, level, *args, **kwargs):

        sep  = kwargs.get(u'sep', u' ')
        end  = kwargs.get(u'end', u'\n')

        html = u''
        text = u''
        for arg in args:
            if isbytestring(arg):
                arg = force_unicode(arg)
            elif not isinstance(arg, unicode):
                arg = as_unicode(arg)
            html += arg + sep
            text += arg + sep

        self.html_log.emit(self.color[level] + escape(html + end) + self.normal)
        self.plain_text_log.emit(text + end)


if __name__ == '__main__':
    import sys

    # compile config.ui to config_ui.py
    # usage: calibre-debug -e config.py config.ui config_ui.py
    # ref. calibre.gui2.__init__.py#build_forms
    if 2 < len(sys.argv):
        import re, cStringIO
        from PyQt5.uic import compileUi

        pat = re.compile(r'''(['"]):/images/([^'"]+)\1''')
        def sub(match):
            ans = 'I(%s%s%s)'%(match.group(1), match.group(2), match.group(1))
            return ans
        transdef_pat = re.compile(r'^\s+_translate\s+=\s+QtCore.QCoreApplication.translate$', flags=re.M)
        transpat = re.compile(r'_translate\s*\(.+?,\s+"(.+?)(?<!\\)"\)', re.DOTALL)

        buf = cStringIO.StringIO()
        compileUi(sys.argv[1], buf)
        dat = buf.getvalue()
        dat = dat.replace('import images_rc', '')
        dat = transdef_pat.sub('', dat)
        dat = transpat.sub(r'_("\1")', dat)
        dat = dat.replace('_("MMM yyyy")', '"MMM yyyy"')
        dat = dat.replace('_("d MMM yyyy")', '"d MMM yyyy"')
        dat = pat.sub(sub, dat)

        dat = dat.replace('self.program = QtWidgets.QPlainTextEdit(Form)',
         'from calibre.gui2.tweak_book.editor.text import TextEdit\n'
         '        self.program = TextEdit(Form)')

        open(sys.argv[2], 'wb').write(dat)

    # else just execute test.
    else:
        from PyQt5.Qt import QApplication
        from calibre.gui2.preferences import test_widget
        app = QApplication([])
        test_widget('Advanced', 'Plugins')
