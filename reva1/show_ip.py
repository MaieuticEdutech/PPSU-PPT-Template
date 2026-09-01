#!/usr/bin/env python3
"""Show this machine's local (LAN) IPv4 address on Windows.

Handy when sharing the PPT Designer site with people on the same network --
your IP is handed out by DHCP and can change after a reboot or reconnect.

Usage:
    python show_ip.py           full report
    python show_ip.py --bare    just the IP address (used by run.bat)
"""
import socket
import sys


def primary_lan_ip():
    """The IPv4 of the interface Windows actually uses to reach the network.

    Opening a UDP socket toward a public address makes the OS choose that
    interface and bind a local address to it. UDP is connectionless, so no
    packets are actually sent -- this works even without internet access, as
    long as a network route exists. This is why it beats simply taking the
    first address from ipconfig, which is often a VirtualBox/WSL adapter.
    """
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except OSError:
        return None
    finally:
        s.close()


def all_ipv4():
    """Every IPv4 address bound to this host, loopback excluded."""
    found = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if not ip.startswith("127."):
                found.add(ip)
    except socket.gaierror:
        pass
    return sorted(found)


def main():
    primary = primary_lan_ip()

    # machine-readable mode: print nothing but the address
    if "--bare" in sys.argv:
        if primary:
            print(primary)
        return

    others = [ip for ip in all_ipv4() if ip != primary]

    print("=" * 46)
    if primary:
        print("Your local IPv4 address:  " + primary)
    else:
        print("Could not determine your IPv4 address.")
        print("Are you connected to a network?")
    print("=" * 46)

    if others:
        print()
        print("Other adapters found: " + ", ".join(others))
        print("(usually VirtualBox / VMware / WSL / Hyper-V -- ignore these)")

    if primary:
        print()
        print("Share these with people on your network:")
        print("  Website:  http://%s:8000/index.html" % primary)
        print("  Backend:  http://%s:5000/" % primary)
        print()
        print("If this IP changed, update this line in frontend/index.html:")
        print("    const BACKEND_HOST = '%s';" % primary)


if __name__ == "__main__":
    main()
