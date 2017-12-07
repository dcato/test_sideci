#!/usr/bin/env python
# vim:fileencoding=UTF-8:ts=4:sw=4:sta:et:sts=4:ai
from __future__ import (unicode_literals, division, absolute_import,
                        print_function)

__license__   = 'GPL v3'
__copyright__ = '2016, D.Cato'
__docformat__ = 'restructuredtext en'

import re
from functools import partial

from PyQt5.Qt import QMenu, QIcon

from calibre.utils.logging import GUILog, ANSIStream
from calibre.gui2 import error_dialog
from calibre.gui2.actions import InterfaceAction, menu_action_unique_name

from calibre_plugins.exec_macro.config import prefs

from calibre_plugins.exec_macro.config import show_config_dialog
from calibre_plugins.exec_macro.config import ConfigWidget


class ExecMacroAction(InterfaceAction):

    name = 'Exec Macro'
    # Create our top-level menu/toolbar action (text, icon_path, tooltip, keyboard shortcut)
    action_spec = ('Exec Macro', None, 'Execute current selected macro.', ())
    action_type = 'current'

    def genesis(self):
        # This method is called once per plugin, do initial setup here
        self.menu = QMenu(self.gui)
        self.actions = {}
        self.rebuild_menu()

        self.qaction.setMenu(self.menu)
        self.qaction.setIcon(get_icons('images/icon.png'))
        self.qaction.triggered.connect(self.execute_current_macro)

    def rebuild_menu(self):
        self.menu.clear()
        self.actions.clear()

        macros = prefs['macros']
        for name in sorted(macros):
            action = self.create_menu_action_unique(self.menu, name,
                tooltip=macros[name]['documentation'],
                image='dot_red.png', shortcut=None, shortcut_name=None,
                triggered=partial(self.execute_macro, name))
            self.actions[name] = action
        self.mark_current_macro()

        self.menu.addSeparator()
        self.create_menu_action_unique(self.menu, _('Manage macros'), tooltip=None,
            image='config.png', shortcut=None, shortcut_name=None,
            triggered=partial(show_config_dialog, self))

        self.delete_orphan_shortcuts()

    def delete_orphan_shortcuts(self):
        prefix = menu_action_unique_name(self, '')
        to_del = {sc for sc in self.gui.keyboard.shortcuts if sc.startswith(prefix)}
        new_sc = {menu_action_unique_name(self, name) for name in self.actions}
        new_sc.add(menu_action_unique_name(self, _('Manage macros')))
        to_del -= new_sc
        for sc in to_del:
            self.gui.keyboard.unregister_shortcut(sc)
        self.gui.keyboard.finalize()

    def mark_current_macro(self):
        for name, action in self.actions.iteritems():
            action.setIconVisibleInMenu(name == prefs['current_macro'])

    def execute_current_macro(self):
        self.execute_macro(prefs['current_macro'])

    def execute_macro(self, name):
        if name != prefs['current_macro']:
            prefs['current_macro'] = name
            self.mark_current_macro()

        macro = prefs['macros'].get(name, None)
        if not macro:
            return

        log = GUILog()
        log.outputs.append(ANSIStream())
        try:
            if macro.get('execfromfile'):
                with open(macro['macrofile'], 'rU') as file:
                    self.execute(file, log)
            elif macro['program']:
                program = macro['program']
                encoding = self.get_encoding(program)
                if encoding:
                    program = program.encode(encoding)
                self.execute(program, log)
        except:
            log.exception(_('Failed to execute macro'))
            error_dialog(self.gui, _('Failed to execute macro'), _(
                'Failed to execute macro, click "Show details" for more information.'),
                 det_msg=log.plain_text, show=True)



    def execute(self, program, log):
        # exec(program, globals().copy(), locals().copy())
        vars = globals().copy()
        vars['self'] = self
        vars['log'] = log
        exec(program, vars)


    def create_menu_action_unique(self, parent_menu, menu_text, image=None, tooltip=None,
        shortcut=None, triggered=None, is_checked=None, shortcut_name=None, unique_name=None):

        '''
        Create a menu action with the specified criteria and action, using the new
        InterfaceAction.create_menu_action() function which ensures that regardless of
        whether a shortcut is specified it will appear in Preferences->Keyboard
        '''
        orig_shortcut = shortcut
        kb = self.gui.keyboard
        if unique_name is None:
            unique_name = menu_text
        if not shortcut == False:
            full_unique_name = menu_action_unique_name(self, unique_name)
            if full_unique_name in kb.shortcuts:
                shortcut = False
            else:
                if shortcut is not None and not shortcut == False:
                    if len(shortcut) == 0:
                        shortcut = None
                    else:
                        shortcut = _(shortcut)

        if shortcut_name is None:
            shortcut_name = menu_text.replace('&','')

        ac = self.create_menu_action(parent_menu, unique_name, menu_text, icon=None, shortcut=shortcut,
            description=tooltip, triggered=triggered, shortcut_name=shortcut_name)
        if shortcut == False and not orig_shortcut == False:
            if ac.calibre_shortcut_unique_name in self.gui.keyboard.shortcuts:
                kb.replace_action(ac.calibre_shortcut_unique_name, ac)
        if image:
            ac.setIcon(QIcon(I(image)))
        if is_checked is not None:
            ac.setCheckable(True)
            if is_checked:
                ac.setChecked(True)

        return ac

    def get_encoding(self, txt):
        decl_re = re.compile(r'^[ \t\f]*#.*coding[:=][ \t]*([-\w.]+)')
        blank_re = re.compile(r'^[ \t\f]*(?:[#\r\n]|$)')

        lines = txt.splitlines()

        if len(lines) < 1:
            return None
        match = decl_re.match(lines[0])
        if match:
            return match.group(1)

        if len(lines) < 2 or (not blank_re.match(lines[0])):
            return None
        match = decl_re.match(lines[1])
        if match:
            return match.group(1)

        return None
