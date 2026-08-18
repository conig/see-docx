/* Drive real compositor pointer events for the GTK smoke tests. */

#define _POSIX_C_SOURCE 200809L

#include <errno.h>
#include <linux/input-event-codes.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <wayland-client.h>

#include "wlr-virtual-pointer-unstable-v1-client-protocol.h"

static struct zwlr_virtual_pointer_manager_v1 *manager;

static uint32_t timestamp_ms(void) {
    struct timespec now;

    if (clock_gettime(CLOCK_MONOTONIC, &now) != 0) {
        perror("clock_gettime");
        exit(EXIT_FAILURE);
    }
    return (uint32_t)(now.tv_sec * 1000ULL + now.tv_nsec / 1000000ULL);
}

static void registry_global(
    void *data,
    struct wl_registry *registry,
    uint32_t name,
    const char *interface,
    uint32_t version
) {
    (void)data;
    if (strcmp(interface, zwlr_virtual_pointer_manager_v1_interface.name) != 0)
        return;
    manager = wl_registry_bind(
        registry,
        name,
        &zwlr_virtual_pointer_manager_v1_interface,
        version < 2 ? version : 2
    );
}

static void registry_global_remove(
    void *data, struct wl_registry *registry, uint32_t name
) {
    (void)data;
    (void)registry;
    (void)name;
}

static const struct wl_registry_listener registry_listener = {
    .global = registry_global,
    .global_remove = registry_global_remove,
};

static void flush_frame(
    struct wl_display *display, struct zwlr_virtual_pointer_v1 *pointer
) {
    zwlr_virtual_pointer_v1_frame(pointer);
    if (wl_display_flush(display) < 0 && errno != EAGAIN) {
        perror("wl_display_flush");
        exit(EXIT_FAILURE);
    }
}

int main(void) {
    struct wl_display *display = wl_display_connect(NULL);
    if (display == NULL) {
        fputs("could not connect to the Wayland display\n", stderr);
        return EXIT_FAILURE;
    }
    struct wl_registry *registry = wl_display_get_registry(display);
    wl_registry_add_listener(registry, &registry_listener, NULL);
    if (wl_display_roundtrip(display) < 0 || manager == NULL) {
        fputs("compositor has no wlr virtual-pointer manager\n", stderr);
        return EXIT_FAILURE;
    }
    struct zwlr_virtual_pointer_v1 *pointer =
        zwlr_virtual_pointer_manager_v1_create_virtual_pointer(manager, NULL);
    if (pointer == NULL || wl_display_roundtrip(display) < 0) {
        fputs("could not create a Wayland virtual pointer\n", stderr);
        return EXIT_FAILURE;
    }

    puts("READY");
    fflush(stdout);
    char line[128];
    while (fgets(line, sizeof(line), stdin) != NULL) {
        double first = 0.0;
        double second = 0.0;
        int discrete = 0;
        unsigned int absolute_x = 0;
        unsigned int absolute_y = 0;
        unsigned int button = 0;
        unsigned int state = 0;
        unsigned int x_extent = 0;
        unsigned int y_extent = 0;
        if (
            sscanf(
                line,
                "motion_absolute %u %u %u %u",
                &absolute_x,
                &absolute_y,
                &x_extent,
                &y_extent
            ) == 4
        ) {
            zwlr_virtual_pointer_v1_motion_absolute(
                pointer,
                timestamp_ms(),
                absolute_x,
                absolute_y,
                x_extent,
                y_extent
            );
            flush_frame(display, pointer);
        } else if (sscanf(line, "motion %lf %lf", &first, &second) == 2) {
            zwlr_virtual_pointer_v1_motion(
                pointer,
                timestamp_ms(),
                wl_fixed_from_double(first),
                wl_fixed_from_double(second)
            );
            flush_frame(display, pointer);
        } else if (sscanf(line, "button %u %u", &button, &state) == 2) {
            zwlr_virtual_pointer_v1_button(
                pointer, timestamp_ms(), button, state
            );
            flush_frame(display, pointer);
        } else if (sscanf(line, "scroll %lf %d", &first, &discrete) == 2) {
            uint32_t now = timestamp_ms();
            zwlr_virtual_pointer_v1_axis_source(
                pointer, WL_POINTER_AXIS_SOURCE_WHEEL
            );
            zwlr_virtual_pointer_v1_axis(
                pointer,
                now,
                WL_POINTER_AXIS_VERTICAL_SCROLL,
                wl_fixed_from_double(first)
            );
            zwlr_virtual_pointer_v1_axis_discrete(
                pointer,
                now,
                WL_POINTER_AXIS_VERTICAL_SCROLL,
                wl_fixed_from_double(first),
                discrete
            );
            flush_frame(display, pointer);
        } else if (strcmp(line, "quit\n") == 0) {
            break;
        } else {
            fprintf(stderr, "invalid pointer command: %s", line);
            return EXIT_FAILURE;
        }
    }

    zwlr_virtual_pointer_v1_destroy(pointer);
    zwlr_virtual_pointer_manager_v1_destroy(manager);
    wl_registry_destroy(registry);
    wl_display_disconnect(display);
    return EXIT_SUCCESS;
}
