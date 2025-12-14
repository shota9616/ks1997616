"""macOS Speech Recognition using native Speech framework."""

import threading
import queue
from typing import Callable, Optional
import objc
from Foundation import NSObject, NSRunLoop, NSDefaultRunLoopMode, NSDate
from AVFoundation import (
    AVAudioEngine,
    AVAudioSession,
    AVAudioSessionCategoryRecord,
    AVAudioSessionModeMeasurement,
)
import Speech
from Speech import (
    SFSpeechRecognizer,
    SFSpeechAudioBufferRecognitionRequest,
    SFSpeechRecognizerAuthorizationStatus,
)


class SpeechRecognitionDelegate(NSObject):
    """Delegate for speech recognizer availability changes."""

    def initWithCallback_(self, callback):
        self = objc.super(SpeechRecognitionDelegate, self).init()
        if self is None:
            return None
        self.callback = callback
        return self

    def speechRecognizer_availabilityDidChange_(self, recognizer, available):
        if self.callback:
            self.callback(available)


class MacOSSpeechRecognizer:
    """Speech recognizer using macOS native Speech framework."""

    def __init__(self, locale: str = "ja-JP"):
        """Initialize the speech recognizer.

        Args:
            locale: Language locale (e.g., "ja-JP" for Japanese, "en-US" for English)
        """
        self.locale = locale
        self._recognizer: Optional[SFSpeechRecognizer] = None
        self._audio_engine: Optional[AVAudioEngine] = None
        self._recognition_request: Optional[SFSpeechAudioBufferRecognitionRequest] = None
        self._recognition_task = None
        self._is_listening = False
        self._result_queue: queue.Queue = queue.Queue()
        self._delegate = None

    def check_authorization(self) -> bool:
        """Check and request speech recognition authorization."""
        status = SFSpeechRecognizer.authorizationStatus()

        if status == SFSpeechRecognizerAuthorizationStatus.authorized:
            return True
        elif status == SFSpeechRecognizerAuthorizationStatus.notDetermined:
            # Request authorization
            auth_event = threading.Event()
            auth_result = [False]

            def auth_handler(status):
                auth_result[0] = (status == SFSpeechRecognizerAuthorizationStatus.authorized)
                auth_event.set()

            SFSpeechRecognizer.requestAuthorization_(auth_handler)
            auth_event.wait(timeout=30)
            return auth_result[0]
        else:
            return False

    def _setup_audio_engine(self):
        """Set up the audio engine for recording."""
        self._audio_engine = AVAudioEngine.alloc().init()

        input_node = self._audio_engine.inputNode()
        recording_format = input_node.outputFormatForBus_(0)

        input_node.installTapOnBus_bufferSize_format_block_(
            0,
            1024,
            recording_format,
            self._audio_buffer_handler
        )

    def _audio_buffer_handler(self, buffer, when):
        """Handle incoming audio buffers."""
        if self._recognition_request:
            self._recognition_request.appendAudioPCMBuffer_(buffer)

    def start_listening(
        self,
        on_result: Optional[Callable[[str, bool], None]] = None,
        on_error: Optional[Callable[[str], None]] = None
    ):
        """Start listening for speech.

        Args:
            on_result: Callback function(text, is_final) called when speech is recognized
            on_error: Callback function(error_message) called on errors
        """
        if self._is_listening:
            return

        if not self.check_authorization():
            if on_error:
                on_error("Speech recognition not authorized")
            return

        # Create recognizer with locale
        locale_obj = objc.lookUpClass('NSLocale').alloc().initWithLocaleIdentifier_(self.locale)
        self._recognizer = SFSpeechRecognizer.alloc().initWithLocale_(locale_obj)

        if not self._recognizer or not self._recognizer.isAvailable():
            if on_error:
                on_error(f"Speech recognizer not available for locale: {self.locale}")
            return

        # Set up audio engine
        self._setup_audio_engine()

        # Create recognition request
        self._recognition_request = SFSpeechAudioBufferRecognitionRequest.alloc().init()
        self._recognition_request.setShouldReportPartialResults_(True)

        # Start recognition task
        def result_handler(result, error):
            if error:
                error_desc = str(error.localizedDescription()) if error else "Unknown error"
                if on_error and "216" not in error_desc:  # Ignore end-of-speech errors
                    on_error(error_desc)
                return

            if result:
                text = result.bestTranscription().formattedString()
                is_final = result.isFinal()

                if on_result:
                    on_result(text, is_final)

                if is_final:
                    self._result_queue.put(text)

        self._recognition_task = self._recognizer.recognitionTaskWithRequest_resultHandler_(
            self._recognition_request,
            result_handler
        )

        # Start audio engine
        self._audio_engine.prepare()
        success, error = self._audio_engine.startAndReturnError_(None)
        if not success:
            if on_error:
                on_error(f"Failed to start audio engine: {error}")
            return

        self._is_listening = True

    def stop_listening(self) -> Optional[str]:
        """Stop listening and return the final recognized text."""
        if not self._is_listening:
            return None

        self._is_listening = False

        # Stop audio engine
        if self._audio_engine:
            self._audio_engine.stop()
            self._audio_engine.inputNode().removeTapOnBus_(0)

        # End recognition request
        if self._recognition_request:
            self._recognition_request.endAudio()

        # Cancel recognition task
        if self._recognition_task:
            self._recognition_task.cancel()

        # Get final result
        try:
            result = self._result_queue.get(timeout=2.0)
            return result
        except queue.Empty:
            return None

    def listen_once(self, timeout: float = 10.0) -> Optional[str]:
        """Listen for a single utterance and return the recognized text.

        Args:
            timeout: Maximum time to wait for speech in seconds

        Returns:
            Recognized text or None if no speech detected
        """
        result_text = [None]
        done_event = threading.Event()

        def on_result(text: str, is_final: bool):
            if is_final:
                result_text[0] = text
                done_event.set()

        def on_error(error: str):
            print(f"Recognition error: {error}")
            done_event.set()

        self.start_listening(on_result=on_result, on_error=on_error)

        # Wait for result or timeout
        done_event.wait(timeout=timeout)
        self.stop_listening()

        return result_text[0]

    @property
    def is_listening(self) -> bool:
        """Check if currently listening."""
        return self._is_listening


class SimpleSpeechInput:
    """Simplified speech input that uses macOS dictation via subprocess."""

    def __init__(self):
        pass

    def listen_with_prompt(self, prompt: str = "話しかけてください...") -> Optional[str]:
        """Show prompt and wait for speech input using a simpler method.

        This method uses a terminal-based approach where the user can:
        1. Press Enter to start recording
        2. Speak
        3. Press Enter again to stop

        For actual dictation, users should use macOS's built-in dictation
        (press Fn twice) while the input prompt is active.
        """
        print(f"\n🎤 {prompt}")
        print("(macOSのDictationを使う場合: Fnキーを2回押して話す)")
        print("(テキスト入力も可能です)")

        try:
            text = input("> ")
            return text.strip() if text.strip() else None
        except (KeyboardInterrupt, EOFError):
            return None


def create_recognizer(use_native: bool = True, locale: str = "ja-JP"):
    """Create a speech recognizer.

    Args:
        use_native: If True, try to use native macOS Speech framework
        locale: Language locale

    Returns:
        Speech recognizer instance
    """
    if use_native:
        try:
            recognizer = MacOSSpeechRecognizer(locale=locale)
            if recognizer.check_authorization():
                return recognizer
        except Exception as e:
            print(f"Native speech recognition unavailable: {e}")

    # Fallback to simple input
    return SimpleSpeechInput()
