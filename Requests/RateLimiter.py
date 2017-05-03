from time import time, sleep

_READ_REQUESTS = 0
_WRITE_REQUESTS = 0

_MAX_WRITE_PER_100 = 100
_MAX_WRITE_DAILY = 40000

_MAX_READ_PER_100 = 100
_MAX_READ_DAILY = 40000

start = time()


def read_request():
    global _READ_REQUESTS
    _READ_REQUESTS += 1
    sleep(1)


def write_request():
    global _WRITE_REQUESTS
    _WRITE_REQUESTS += 1
    sleep(1)
