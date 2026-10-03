import logging
import os
import tempfile
import unittest

import tests  # noqa: F401
from office_plugin.core.dispatch import Router
from office_plugin.core.events import EventBus
from office_plugin.core.feature import Feature
from office_plugin.core.registry import Registry
from office_plugin.core.services import ServiceContainer
from office_plugin.core.settings import Settings

LOG = logging.getLogger("office_plugin_test")
LOG.addHandler(logging.NullHandler())


class EventBusTests(unittest.TestCase):
    def test_publish_subscribe_unsubscribe(self):
        bus, seen = EventBus(LOG), []
        token = bus.subscribe("x", lambda **kw: seen.append(kw))
        bus.publish("x", a=1)
        bus.unsubscribe(token)
        bus.publish("x", a=2)
        self.assertEqual(seen, [{"a": 1}])

    def test_failing_subscriber_does_not_block_others(self):
        bus, seen = EventBus(LOG), []
        bus.subscribe("x", lambda **kw: 1 / 0)
        bus.subscribe("x", lambda **kw: seen.append(1))
        bus.publish("x")
        self.assertEqual(seen, [1])


class ServicesTests(unittest.TestCase):
    def test_lazy_and_cached(self):
        s, calls = ServiceContainer(), []
        s.register("a", lambda: calls.append(1) or object())
        self.assertTrue(s.has("a"))
        self.assertIs(s.get("a"), s.get("a"))
        self.assertEqual(len(calls), 1)
        with self.assertRaises(KeyError):
            s.get("nope")


class SettingsTests(unittest.TestCase):
    def test_roundtrip_and_corrupt_file(self):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "sub", "settings.json")
            s = Settings(path)
            self.assertTrue(s.feature_enabled("navigation"))
            s.set_feature_enabled("navigation", False)
            s.set_option("navigation", "max_entries", 5)
            s.save()
            s2 = Settings(path)
            self.assertFalse(s2.feature_enabled("navigation"))
            self.assertEqual(s2.option("navigation", "max_entries"), 5)
            with open(path, "w") as fh:
                fh.write("{not json")
            self.assertTrue(Settings(path).feature_enabled("navigation"))  # falls back to defaults


class RegistryAndRouterTests(unittest.TestCase):
    def make(self, settings=None):
        services = ServiceContainer()
        for name in ("positions", "doc_utils", "ui", "log"):
            services.register_instance(name, object())
        return Registry(None, services, EventBus(LOG), settings or Settings(None), LOG)

    def test_loads_all_shipped_features(self):
        reg = self.make()
        reg.load()
        self.assertEqual(reg.failed, {})
        self.assertEqual(set(reg.features), {"navigation", "last_edit", "calc_visible", "calc_dedupe", "calc_setops", "calc_fill_blanks",
                         "calc_remove_blanks", "calc_highlight", "calc_transpose",
                         "calc_sequence", "calc_missing_seq", "calc_gantt",
                         "calc_decimals", "calc_numwords", "calc_units", "calc_functions"})
        reg.stop()

    def test_disabled_feature_is_not_started(self):
        settings = Settings(None)
        settings.set_feature_enabled("navigation", False)
        reg = self.make(settings)
        reg.load()
        self.assertNotIn("navigation", reg.features)
        self.assertIn("navigation", reg.manifests)
        reg.stop()

    def test_missing_service_isolates_the_failure(self):
        reg = self.make()
        reg.services = ServiceContainer()      # nearly no services
        reg.services.register_instance("doc_utils", object())
        reg.load()
        self.assertIn("navigation", reg.failed)              # needs positions
        self.assertIn("calc_dedupe", reg.failed)             # needs ui
        self.assertIn("calc_visible", reg.failed)            # needs ui
        self.assertEqual(reg.features, {})

    def test_router_resolves_and_gates_by_app(self):
        reg = self.make()
        reg.load()
        router = Router(reg, LOG)
        self.assertIsNotNone(router.resolve("vnd.officeplugin:navigation.Back"))
        self.assertIsNone(router.resolve("vnd.officeplugin:navigation.Nope"))
        self.assertIsNone(router.resolve("vnd.officeplugin:ghost.Back"))
        self.assertIsNone(router.resolve(".uno:Save"))
        self.assertFalse(router.is_enabled("vnd.officeplugin:calc_dedupe.Run", None, "writer"))
        self.assertTrue(router.is_enabled("vnd.officeplugin:calc_dedupe.Run", None, "calc"))
        reg.stop()

    def test_router_isolates_command_errors(self):
        class Boom(Feature):
            def execute(self, command, frame, args):
                raise RuntimeError("boom")

        reg = self.make()
        reg.features["boom"] = Boom(None, reg.services, reg.events, reg.settings,
                                    {"id": "boom", "commands": [{"name": "Go"}]}, LOG)
        self.assertFalse(Router(reg, LOG).execute("vnd.officeplugin:boom.Go", None))

    def test_feature_stop_unsubscribes(self):
        class Listener(Feature):
            def start(self):
                self.subscribe("ping", lambda **kw: self.hits.append(1))

        bus = EventBus(LOG)
        f = Listener(None, None, bus, Settings(None), {"id": "l", "commands": []}, LOG)
        f.hits = []
        f.start()
        bus.publish("ping")
        f.stop()
        bus.publish("ping")
        self.assertEqual(f.hits, [1])


if __name__ == "__main__":
    unittest.main()
